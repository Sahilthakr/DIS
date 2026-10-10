VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form hemcbcreport 
   Caption         =   "Complete Blood Count Report"
   ClientHeight    =   10305
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13620
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
   ScaleHeight     =   10305
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Surface Markers"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      MaskColor       =   &H8000000A&
      TabIndex        =   246
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Frame Frame11 
      Caption         =   "DNA"
      Height          =   855
      Left            =   120
      TabIndex        =   241
      Top             =   7080
      Width           =   9375
      Begin VB.TextBox TXT_DNAINDEX 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   243
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TXT_DNACLOIDY 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   3360
         TabIndex        =   242
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label8 
         Caption         =   "INDEX"
         Height          =   255
         Left            =   120
         TabIndex        =   244
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txt_SAVE 
      Height          =   285
      Left            =   4560
      TabIndex        =   237
      Top             =   10920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txt_VIEW 
      Height          =   285
      Left            =   2880
      TabIndex        =   236
      Top             =   10920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt_NEW 
      Height          =   285
      Left            =   1080
      TabIndex        =   235
      Top             =   10680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport CrystalReport4 
      Left            =   720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport3 
      Left            =   720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\charulata\dis\REPORT\DISCPDET.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cmb_enterby 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "hemcbcreport.frx":0000
      Left            =   7800
      List            =   "hemcbcreport.frx":0002
      TabIndex        =   60
      Top             =   480
      Width           =   2295
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame COMMAND 
      Height          =   735
      Left            =   120
      TabIndex        =   46
      Top             =   9360
      Width           =   11775
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         Height          =   375
         Left            =   7200
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   10440
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   8760
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   225
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5880
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4440
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   3120
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1800
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Default         =   -1  'True
         Height          =   375
         Left            =   360
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame12 
      Height          =   1335
      Left            =   120
      TabIndex        =   44
      Top             =   7920
      Width           =   11775
      Begin VB.TextBox txt_1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   2000
         TabIndex        =   239
         Top             =   240
         Visible         =   0   'False
         Width           =   7065
      End
      Begin VB.CommandButton cmd_mdc 
         Caption         =   "MDC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         MousePointer    =   99  'Custom
         TabIndex        =   68
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox txt_comments 
         Height          =   315
         Left            =   1080
         TabIndex        =   67
         Top             =   180
         Width           =   7905
      End
      Begin VB.ComboBox cmb_finalby 
         Height          =   315
         Left            =   9180
         TabIndex        =   7
         Top             =   900
         Width           =   2415
      End
      Begin VB.ComboBox CMB_PROVISIONAL 
         Height          =   315
         Left            =   9240
         TabIndex        =   66
         Top             =   900
         Width           =   1935
      End
      Begin VB.TextBox txt_addendum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   64
         Top             =   630
         Width           =   2535
      End
      Begin VB.ComboBox regcbo 
         Height          =   315
         Left            =   1170
         TabIndex        =   62
         Top             =   900
         Width           =   2055
      End
      Begin VB.ComboBox cmb_addendumby 
         Height          =   315
         Left            =   4860
         TabIndex        =   58
         Top             =   990
         Width           =   2295
      End
      Begin VB.TextBox txt_impression 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   540
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   90
         TabIndex        =   63
         Top             =   270
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Registrar"
         Height          =   255
         Left            =   180
         TabIndex        =   61
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Addendum By"
         Height          =   255
         Left            =   3510
         TabIndex        =   57
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Addendum"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3600
         TabIndex        =   56
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Impression"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   90
         TabIndex        =   55
         Top             =   630
         Width           =   915
      End
      Begin VB.Label lbl_finby 
         AutoSize        =   -1  'True
         Caption         =   "To be Finalized by"
         Height          =   195
         Left            =   7440
         TabIndex        =   45
         Top             =   990
         Width           =   1575
      End
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   11520
      TabIndex        =   42
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   11520
      TabIndex        =   41
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame common 
      Height          =   2535
      Left            =   0
      TabIndex        =   5
      Top             =   840
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
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox CMB_PREPATHNO 
         BackColor       =   &H008080FF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9120
         TabIndex        =   53
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   0
         TabIndex        =   23
         Top             =   720
         Width           =   1335
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmb_reportdoc 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   0
         TabIndex        =   22
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
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6720
         TabIndex        =   21
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txt_modcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8040
         TabIndex        =   18
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   17
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   13
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   11
         Top             =   1080
         Width           =   6015
      End
      Begin VB.TextBox TXT_PATH 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   10
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
      Begin VB.TextBox txt_deptcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   9
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
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Prv. ReqNo."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   54
         Top             =   1800
         Width           =   1065
      End
      Begin VB.Label lbl_prodis 
         AutoSize        =   -1  'True
         Caption         =   "Provisional Diagnosis"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4800
         TabIndex        =   40
         Top             =   2160
         Width           =   1830
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   2160
         Width           =   2130
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5400
         TabIndex        =   38
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_ward 
         Caption         =   "Ward  BED"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         TabIndex        =   37
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9480
         TabIndex        =   35
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   34
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5040
         TabIndex        =   33
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   32
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1680
         TabIndex        =   31
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "Hem No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9360
         TabIndex        =   30
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   29
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   28
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition Placed to"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   1830
      End
   End
   Begin Crystal.CrystalReport CrystalReport5 
      Left            =   1200
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   0
      TabIndex        =   69
      Top             =   3480
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   192
      TabCaption(0)   =   "C.B.C (Counter)"
      TabPicture(0)   =   "hemcbcreport.frx":0004
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "C.B.C + PBS"
      TabPicture(1)   =   "hemcbcreport.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Coagulation Profile"
      TabPicture(2)   =   "hemcbcreport.frx":003C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Cytochemistry "
      TabPicture(3)   =   "hemcbcreport.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Suface Marker"
      TabPicture(4)   =   "hemcbcreport.frx":0074
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "ESR"
      TabPicture(5)   =   "hemcbcreport.frx":0090
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label29"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "txt_esr"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      Begin VB.TextBox txt_esr 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   -71280
         TabIndex        =   234
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Frame Frame8 
         Caption         =   "Platelets"
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   -70200
         TabIndex        =   225
         Top             =   3360
         Width           =   6855
         Begin VB.TextBox TXT_ADEQUACY 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4320
            TabIndex        =   226
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label LBL_ADEQUACY 
            AutoSize        =   -1  'True
            Caption         =   "Adequacy"
            Height          =   195
            Left            =   3360
            TabIndex        =   228
            Top             =   240
            Width           =   855
         End
         Begin VB.Label LBL_MORPHOLOGY 
            AutoSize        =   -1  'True
            Caption         =   "Morphology"
            Height          =   195
            Left            =   120
            TabIndex        =   227
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "RBC"
         ForeColor       =   &H00800000&
         Height          =   3015
         Left            =   -68520
         TabIndex        =   215
         Top             =   360
         Width           =   5175
         Begin VB.TextBox txt_cbc_comment 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            TabIndex        =   238
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox txt_othertype 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   190
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox TXT_LP 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   188
            Top             =   600
            Width           =   855
         End
         Begin VB.ComboBox CMB_MORPHO 
            Height          =   315
            ItemData        =   "hemcbcreport.frx":00AC
            Left            =   3120
            List            =   "hemcbcreport.frx":00BC
            TabIndex        =   194
            Top             =   120
            Width           =   1935
         End
         Begin VB.TextBox txt_malarialparasite 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            MaxLength       =   250
            ScrollBars      =   1  'Horizontal
            TabIndex        =   193
            Top             =   2595
            Width           =   3375
         End
         Begin VB.TextBox txt_reticcount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4320
            TabIndex        =   196
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txt_rbcmorphology 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            MaxLength       =   250
            TabIndex        =   187
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox TXT_OTHERCELLSRBC 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1920
            TabIndex        =   191
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3120
            MaxLength       =   50
            TabIndex        =   216
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbmpo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "hemcbcreport.frx":0118
            Left            =   1320
            List            =   "hemcbcreport.frx":0125
            TabIndex        =   192
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txt_empc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4320
            MaxLength       =   50
            TabIndex        =   195
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txt_rpi 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4320
            TabIndex        =   198
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txt_crc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4320
            TabIndex        =   197
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox TXT_MORPHOLOGY 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4200
            MaxLength       =   100
            TabIndex        =   199
            Top             =   240
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.TextBox txt_rc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   189
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label33 
            Caption         =   "Lap Score                     Estimated Manual                                            Platelet Count"
            Height          =   375
            Left            =   120
            TabIndex        =   231
            Top             =   600
            Width           =   4095
         End
         Begin VB.Label lbl_malarialparasite 
            Caption         =   "Malarial Parasite"
            Height          =   435
            Left            =   120
            TabIndex        =   224
            Top             =   2520
            Width           =   705
         End
         Begin VB.Label lbl_reticcount 
            AutoSize        =   -1  'True
            Caption         =   "Retic Count [%]"
            Height          =   195
            Left            =   2760
            TabIndex        =   223
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label LBL_RBCMORPHOLOGY 
            Caption         =   "RBC Morphology"
            ForeColor       =   &H000040C0&
            Height          =   435
            Left            =   120
            TabIndex        =   222
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label LBL_OTHERRBC 
            AutoSize        =   -1  'True
            Caption         =   "Other Cells"
            Height          =   195
            Left            =   120
            TabIndex        =   221
            Top             =   1755
            Width           =   945
         End
         Begin VB.Label LBL_PLATMORPHO 
            Caption         =   "Platelets  Morphology"
            ForeColor       =   &H000040C0&
            Height          =   435
            Left            =   2280
            TabIndex        =   220
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "MPO "
            Height          =   195
            Left            =   120
            TabIndex        =   219
            Top             =   2160
            Width           =   480
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "CRC  RPI"
            Height          =   615
            Left            =   2760
            TabIndex        =   218
            Top             =   1800
            Width           =   375
            WordWrap        =   -1  'True
         End
         Begin VB.Label LBL_RC 
            Caption         =   "Recticulocyte Count"
            Height          =   435
            Left            =   120
            TabIndex        =   217
            Top             =   1200
            Width           =   1620
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "WBC (%)"
         ForeColor       =   &H00800000&
         Height          =   3135
         Left            =   -75000
         TabIndex        =   168
         Top             =   360
         Width           =   6495
         Begin VB.TextBox txt_prolymphocytes 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   232
            Top             =   840
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txt_prolympho 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5880
            TabIndex        =   181
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox TXT_PROMYELOCYTESWBC 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   170
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txt_othercells 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5640
            TabIndex        =   184
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox txt_hairycells 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5640
            TabIndex        =   183
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txt_Sezarycells 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5640
            TabIndex        =   182
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txt_promonocyte 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   171
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox txt_alymphocytes 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3600
            TabIndex        =   180
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox txt_plasmacells 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3600
            TabIndex        =   179
            Top             =   2280
            Width           =   855
         End
         Begin VB.TextBox txt_blasts 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   169
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txt_myelocytes 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   172
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox txt_metamyelocytes 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   173
            Top             =   2280
            Width           =   855
         End
         Begin VB.TextBox txt_mbasophills 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3600
            TabIndex        =   178
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox txt_meosinophils 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3600
            TabIndex        =   177
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox txt_mmonocytes 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3600
            TabIndex        =   176
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txt_mlymphocytes 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3600
            TabIndex        =   175
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txt_mneutrophils 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   174
            Top             =   2760
            Width           =   855
         End
         Begin VB.TextBox TXT_NORMOBLASTS 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5640
            TabIndex        =   185
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label32 
            Caption         =   "Prolymphocytes"
            Height          =   255
            Left            =   4560
            TabIndex        =   230
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label LBL_PROMYELOCYTESWBC 
            AutoSize        =   -1  'True
            Caption         =   "Promyelocytes"
            Height          =   195
            Left            =   75
            TabIndex        =   214
            Top             =   810
            Width           =   1230
         End
         Begin VB.Label lbl_othercells 
            AutoSize        =   -1  'True
            Caption         =   "Other Cells"
            Height          =   195
            Left            =   4560
            TabIndex        =   213
            Top             =   1800
            Width           =   945
         End
         Begin VB.Label lbl_hairycells 
            Caption         =   "Hairy Cells"
            Height          =   255
            Left            =   4560
            TabIndex        =   212
            Top             =   1395
            Width           =   975
         End
         Begin VB.Label lbl_sezarycells 
            AutoSize        =   -1  'True
            Caption         =   "Sezary Cells"
            Height          =   195
            Left            =   4560
            TabIndex        =   211
            Top             =   825
            Width           =   1050
         End
         Begin VB.Label lbl_promonocyte 
            AutoSize        =   -1  'True
            Caption         =   "Promonocyte"
            Height          =   195
            Left            =   90
            TabIndex        =   210
            Top             =   1320
            Width           =   1110
         End
         Begin VB.Label lbl_matyplicallymphocytes 
            Caption         =   "Atypical Lymphocytes"
            Height          =   435
            Left            =   2400
            TabIndex        =   209
            Top             =   2640
            Width           =   1125
         End
         Begin VB.Label lbl_blasts 
            AutoSize        =   -1  'True
            Caption         =   "Blasts"
            Height          =   195
            Left            =   75
            TabIndex        =   208
            Top             =   360
            Width           =   525
         End
         Begin VB.Label lbl_myelocytes 
            AutoSize        =   -1  'True
            Caption         =   "Myelocytes"
            Height          =   195
            Left            =   75
            TabIndex        =   207
            Top             =   1860
            Width           =   960
         End
         Begin VB.Label lbl_metamyelocytes 
            AutoSize        =   -1  'True
            Caption         =   "Metamyelocytes"
            Height          =   195
            Left            =   75
            TabIndex        =   206
            Top             =   2310
            Width           =   1365
         End
         Begin VB.Label lbl_dmneutrophils 
            AutoSize        =   -1  'True
            Caption         =   "Neutrophils"
            Height          =   195
            Left            =   75
            TabIndex        =   205
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label lbl_mlymphocytes 
            AutoSize        =   -1  'True
            Caption         =   "Lymphocytes"
            Height          =   195
            Left            =   2430
            TabIndex        =   204
            Top             =   360
            Width           =   1110
         End
         Begin VB.Label lbl_mmonocytes 
            AutoSize        =   -1  'True
            Caption         =   "Monocytes"
            Height          =   195
            Left            =   2430
            TabIndex        =   203
            Top             =   840
            Width           =   930
         End
         Begin VB.Label LBL_EOSINOPHILS1 
            AutoSize        =   -1  'True
            Caption         =   "Eosinophils"
            Height          =   195
            Left            =   2430
            TabIndex        =   202
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lbl_mbasophills 
            AutoSize        =   -1  'True
            Caption         =   "Basophills"
            Height          =   195
            Left            =   2430
            TabIndex        =   201
            Top             =   1800
            Width           =   870
         End
         Begin VB.Label lbl_plasmacells 
            AutoSize        =   -1  'True
            Caption         =   "Plasma Cells"
            Height          =   195
            Left            =   2430
            TabIndex        =   200
            Top             =   2280
            Width           =   1080
         End
         Begin VB.Label lbl_normoblastsper100wbc 
            Caption         =   "Normoblasts / 100 WBC"
            Height          =   435
            Left            =   4560
            TabIndex        =   186
            Top             =   2160
            Width           =   1110
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3135
         Left            =   120
         TabIndex        =   141
         Top             =   360
         Width           =   11520
         Begin VB.ComboBox CMB_range 
            Height          =   315
            ItemData        =   "hemcbcreport.frx":013F
            Left            =   5160
            List            =   "hemcbcreport.frx":014F
            TabIndex        =   245
            Top             =   2280
            Width           =   1215
         End
         Begin VB.ComboBox cmb_machname 
            Height          =   315
            ItemData        =   "hemcbcreport.frx":0164
            Left            =   7320
            List            =   "hemcbcreport.frx":0177
            TabIndex        =   240
            Top             =   2640
            Width           =   1335
         End
         Begin VB.TextBox txt_ptctl 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   151
            Text            =   "10.5"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox TXT_OTHERS 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2400
            TabIndex        =   150
            Top             =   2640
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox TXT_TT 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2400
            TabIndex        =   147
            Top             =   1530
            Width           =   855
         End
         Begin VB.TextBox TXT_FDPDMIER 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2430
            TabIndex        =   149
            Top             =   2280
            Width           =   2535
         End
         Begin VB.TextBox TXT_FIBRINOGEN 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2400
            TabIndex        =   148
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox txt_inr 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2400
            TabIndex        =   145
            Top             =   810
            Width           =   855
         End
         Begin VB.TextBox TXT_PT 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   2400
            TabIndex        =   144
            Top             =   450
            Width           =   855
         End
         Begin VB.TextBox TXT_PTTK 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2400
            TabIndex        =   146
            Top             =   1170
            Width           =   855
         End
         Begin VB.TextBox txt_apttctl 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   143
            Text            =   "29.5"
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txt_ttctl 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3360
            TabIndex        =   142
            Top             =   1530
            Width           =   855
         End
         Begin VB.Label LBL_OTHERSCP 
            AutoSize        =   -1  'True
            Caption         =   "Others"
            Height          =   195
            Left            =   1320
            TabIndex        =   167
            Top             =   2640
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label LBL_TT 
            AutoSize        =   -1  'True
            Caption         =   "TT"
            Height          =   195
            Left            =   1320
            TabIndex        =   166
            Top             =   1530
            Width           =   255
         End
         Begin VB.Label LBL_FDPDDIMER 
            AutoSize        =   -1  'True
            Caption         =   "D-DIMER"
            Height          =   195
            Left            =   1320
            TabIndex        =   165
            Top             =   2280
            Width           =   810
         End
         Begin VB.Label LBL_FIBRINOGEN 
            AutoSize        =   -1  'True
            Caption         =   "Fibrinogen"
            Height          =   195
            Left            =   1320
            TabIndex        =   164
            Top             =   1920
            Width           =   900
         End
         Begin VB.Label LBL_BT 
            AutoSize        =   -1  'True
            Caption         =   "INR"
            Height          =   195
            Left            =   1320
            TabIndex        =   163
            Top             =   810
            Width           =   345
         End
         Begin VB.Label LBL_PT 
            AutoSize        =   -1  'True
            Caption         =   "PT"
            Height          =   195
            Left            =   1350
            TabIndex        =   162
            Top             =   450
            Width           =   255
         End
         Begin VB.Label LBL_PTTK 
            AutoSize        =   -1  'True
            Caption         =   "APTT"
            Height          =   195
            Left            =   1320
            TabIndex        =   161
            Top             =   1170
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "TEST"
            Height          =   255
            Left            =   2520
            TabIndex        =   160
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "MEAN NORMAL VALUE"
            Height          =   255
            Left            =   3330
            TabIndex        =   159
            Top             =   240
            Width           =   2145
         End
         Begin VB.Label Label10 
            Caption         =   "(Range: 12.7 to 15.0 Sec.)"
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   6480
            TabIndex        =   158
            Top             =   600
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label Label11 
            Caption         =   "(Range: 22.6 to 35 Sec.)"
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   6480
            TabIndex        =   157
            Top             =   1440
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label12 
            Caption         =   "(Range: 05 to15 Sec.)"
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   6480
            TabIndex        =   156
            Top             =   1800
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label13 
            Caption         =   "(Range: < 200 ng/ml.)"
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   6480
            TabIndex        =   155
            Top             =   2160
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label14 
            Caption         =   "(Range: 146 to 380 mg/dL)"
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   6480
            TabIndex        =   154
            Top             =   2400
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label Label16 
            Caption         =   "RANGES"
            Height          =   255
            Left            =   6960
            TabIndex        =   153
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label26 
            Caption         =   "Select Machine Name"
            Height          =   255
            Left            =   5160
            TabIndex        =   152
            Top             =   2640
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cytochemistry Report on Peripheral Blood"
         Height          =   3015
         Left            =   -74880
         TabIndex        =   123
         Top             =   360
         Width           =   11295
         Begin VB.TextBox TXT_MPO 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3600
            TabIndex        =   136
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox TXT_PAS 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   135
            Top             =   2595
            Visible         =   0   'False
            Width           =   9615
         End
         Begin VB.TextBox TXT_NSE 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3600
            TabIndex        =   134
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox TXT_TRAP 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   133
            Top             =   2160
            Width           =   9615
         End
         Begin VB.ComboBox cmb_nse 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "hemcbcreport.frx":019E
            Left            =   1320
            List            =   "hemcbcreport.frx":01A0
            TabIndex        =   132
            Top             =   1680
            Width           =   2175
         End
         Begin VB.ComboBox cmb_mpo 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "hemcbcreport.frx":01A2
            Left            =   1320
            List            =   "hemcbcreport.frx":01A4
            TabIndex        =   131
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Frame Frame10 
            Height          =   855
            Left            =   240
            TabIndex        =   124
            Top             =   240
            Width           =   6975
            Begin VB.TextBox txt_lapscore 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1080
               TabIndex        =   126
               Top             =   480
               Width           =   2175
            End
            Begin VB.TextBox txt_lapctl 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3360
               TabIndex        =   125
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "LAP Score"
               Height          =   195
               Left            =   120
               TabIndex        =   130
               Top             =   480
               Width           =   915
            End
            Begin VB.Label Label20 
               Caption         =   "TEST"
               Height          =   255
               Left            =   1920
               TabIndex        =   129
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label21 
               Caption         =   "CONTROL"
               Height          =   255
               Left            =   3600
               TabIndex        =   128
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label17 
               Caption         =   "(Range: 30 to 100)"
               ForeColor       =   &H000040C0&
               Height          =   255
               Left            =   5160
               TabIndex        =   127
               Top             =   480
               Width           =   1695
            End
         End
         Begin VB.Label lbl_mpo 
            AutoSize        =   -1  'True
            Caption         =   "MPO"
            Height          =   195
            Left            =   720
            TabIndex        =   140
            Top             =   1320
            Width           =   420
         End
         Begin VB.Label LBL_PAS 
            AutoSize        =   -1  'True
            Caption         =   "PAS"
            Height          =   195
            Left            =   765
            TabIndex        =   139
            Top             =   2640
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LBL_NSE 
            AutoSize        =   -1  'True
            Caption         =   "NSE"
            Height          =   195
            Left            =   750
            TabIndex        =   138
            Top             =   1680
            Width           =   390
         End
         Begin VB.Label LBL_TRAP 
            AutoSize        =   -1  'True
            Caption         =   "TRAP"
            Height          =   195
            Left            =   630
            TabIndex        =   137
            Top             =   2280
            Width           =   510
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   84
         Top             =   360
         Width           =   7455
         Begin VB.CommandButton Command1 
            Caption         =   "Clear"
            Height          =   375
            Left            =   240
            TabIndex        =   118
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox TXT_RBC 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   86
            Top             =   840
            Width           =   855
         End
         Begin VB.Frame Frame9 
            Caption         =   "Differential Leukocyte  Count ( %) And Absolute"
            ForeColor       =   &H00800000&
            Height          =   3015
            Left            =   2760
            TabIndex        =   100
            Top             =   0
            Width           =   4695
            Begin VB.TextBox txt_ntphabs 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               TabIndex        =   229
               Top             =   360
               Width           =   855
            End
            Begin VB.TextBox txt_eosinophils 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1440
               TabIndex        =   92
               Top             =   1800
               Width           =   855
            End
            Begin VB.TextBox txt_basophils 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1440
               TabIndex        =   93
               Top             =   2280
               Width           =   855
            End
            Begin VB.TextBox txt_immaturegranulocytes 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2160
               TabIndex        =   106
               Top             =   2640
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txt_neutrophils 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1440
               TabIndex        =   89
               Top             =   360
               Width           =   855
            End
            Begin VB.TextBox txt_lymphocytes 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1440
               TabIndex        =   90
               Top             =   840
               Width           =   855
            End
            Begin VB.TextBox txt_monocytes 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1440
               TabIndex        =   91
               Top             =   1320
               Width           =   855
            End
            Begin VB.TextBox txt_lymabs 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               TabIndex        =   105
               Top             =   840
               Width           =   855
            End
            Begin VB.TextBox txt_mncyteabs 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               TabIndex        =   104
               Top             =   1320
               Width           =   855
            End
            Begin VB.TextBox txt_esnabs 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               TabIndex        =   103
               Top             =   1800
               Width           =   855
            End
            Begin VB.TextBox txt_bsphabs 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3120
               TabIndex        =   102
               Top             =   2280
               Width           =   855
            End
            Begin VB.Label lbl_eosinophils 
               AutoSize        =   -1  'True
               Caption         =   "Eosinophils %"
               Height          =   195
               Left            =   135
               TabIndex        =   117
               Top             =   1800
               Width           =   1290
            End
            Begin VB.Label lbl_basophils 
               AutoSize        =   -1  'True
               Caption         =   "Basophils %"
               Height          =   195
               Left            =   120
               TabIndex        =   116
               Top             =   2280
               Width           =   1020
            End
            Begin VB.Label lbl_immgranulocytes 
               AutoSize        =   -1  'True
               Caption         =   "Immature Granulocytes"
               Height          =   195
               Left            =   120
               TabIndex        =   115
               Top             =   2640
               Visible         =   0   'False
               Width           =   1950
            End
            Begin VB.Label lbl_neutrophils 
               AutoSize        =   -1  'True
               Caption         =   "Neutrophils %"
               Height          =   195
               Left            =   120
               TabIndex        =   114
               Top             =   360
               Width           =   1170
            End
            Begin VB.Label lbl_lymphocytes 
               Caption         =   "Lymphocytes %"
               Height          =   255
               Left            =   120
               TabIndex        =   113
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label lbl_monocytes 
               AutoSize        =   -1  'True
               Caption         =   "Monocytes %"
               Height          =   195
               Left            =   120
               TabIndex        =   112
               Top             =   1320
               Width           =   1125
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Abs"
               Height          =   195
               Left            =   2520
               TabIndex        =   111
               Top             =   1320
               Width           =   330
            End
            Begin VB.Label Label22 
               Caption         =   "Abs"
               Height          =   255
               Left            =   2520
               TabIndex        =   110
               Top             =   840
               Width           =   495
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Abs"
               Height          =   195
               Left            =   2520
               TabIndex        =   109
               Top             =   360
               Width           =   330
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Abs"
               Height          =   195
               Left            =   2520
               TabIndex        =   108
               Top             =   2280
               Width           =   330
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Abs"
               Height          =   195
               Left            =   2535
               TabIndex        =   107
               Top             =   1800
               Width           =   330
            End
         End
         Begin VB.TextBox TXT_HB 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   85
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox TXT_TWBC 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   87
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox txt_platelets 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   88
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label LBL_RBC 
            AutoSize        =   -1  'True
            Caption         =   "RBC                 10 e 12/L"
            Height          =   195
            Left            =   480
            TabIndex        =   122
            Top             =   960
            Width           =   2250
         End
         Begin VB.Label LBL_HB 
            AutoSize        =   -1  'True
            Caption         =   "Hb                  g/dl"
            Height          =   195
            Left            =   600
            TabIndex        =   121
            Top             =   480
            Width           =   1680
         End
         Begin VB.Label lbl_twbc 
            AutoSize        =   -1  'True
            Caption         =   "TLC                  10 e 9/L"
            Height          =   195
            Left            =   480
            TabIndex        =   120
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label LBL_PLATLETS 
            AutoSize        =   -1  'True
            Caption         =   "Platelets                 10 e 9/L"
            Height          =   195
            Left            =   120
            TabIndex        =   119
            Top             =   1920
            Width           =   2505
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "RBC Indices"
         ForeColor       =   &H00800000&
         Height          =   3015
         Left            =   -67320
         TabIndex        =   70
         Top             =   360
         Width           =   3975
         Begin VB.TextBox TXT_RDW 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   98
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox TXT_ABNORMAL 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   101
            Top             =   2640
            Width           =   1815
         End
         Begin VB.TextBox TXT_MPV 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   99
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox TXT_MCHC 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   97
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox TXT_MCH 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   96
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox TXT_MCV 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   95
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox TXT_HCT 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   94
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl_rwd 
            AutoSize        =   -1  'True
            Caption         =   "RDW"
            Height          =   195
            Left            =   720
            TabIndex        =   83
            Top             =   1800
            Width           =   465
         End
         Begin VB.Label LBL_RWDR 
            AutoSize        =   -1  'True
            Caption         =   "%CV "
            Height          =   195
            Left            =   2520
            TabIndex        =   82
            Top             =   1800
            Width           =   450
         End
         Begin VB.Label LBL_ABNORMAL 
            AutoSize        =   -1  'True
            Caption         =   "Abnormal "
            Height          =   195
            Left            =   240
            TabIndex        =   81
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label LBL_MPVR 
            AutoSize        =   -1  'True
            Caption         =   "fl     "
            Height          =   195
            Left            =   2520
            TabIndex        =   80
            Top             =   2160
            Width           =   420
         End
         Begin VB.Label LBL_MCHCR 
            AutoSize        =   -1  'True
            Caption         =   "g/dl   "
            Height          =   195
            Left            =   2520
            TabIndex        =   79
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label LBL_MCHR 
            AutoSize        =   -1  'True
            Caption         =   "pg  "
            Height          =   195
            Left            =   2520
            TabIndex        =   78
            Top             =   1080
            Width           =   345
         End
         Begin VB.Label LBL_MCVR 
            AutoSize        =   -1  'True
            Caption         =   "fl     "
            Height          =   195
            Left            =   2520
            TabIndex        =   77
            Top             =   720
            Width           =   420
         End
         Begin VB.Label LBL_HCTR 
            AutoSize        =   -1  'True
            Caption         =   "%    "
            Height          =   195
            Left            =   2520
            TabIndex        =   76
            Top             =   360
            Width           =   390
         End
         Begin VB.Label lbl_mpv 
            AutoSize        =   -1  'True
            Caption         =   "MPV"
            Height          =   195
            Left            =   720
            TabIndex        =   75
            Top             =   2160
            Width           =   405
         End
         Begin VB.Label lbl_mchc 
            AutoSize        =   -1  'True
            Caption         =   "MCHC"
            Height          =   195
            Left            =   720
            TabIndex        =   74
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label lbl_mch 
            AutoSize        =   -1  'True
            Caption         =   "MCH"
            Height          =   195
            Left            =   720
            TabIndex        =   73
            Top             =   1080
            Width           =   420
         End
         Begin VB.Label LBL_MCV 
            AutoSize        =   -1  'True
            Caption         =   "MCV"
            Height          =   195
            Left            =   720
            TabIndex        =   72
            Top             =   720
            Width           =   405
         End
         Begin VB.Label lbl_hct 
            AutoSize        =   -1  'True
            Caption         =   "Hct"
            Height          =   195
            Left            =   720
            TabIndex        =   71
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.Label Label29 
         Caption         =   "ESR                             mm at the end of one hour."
         Height          =   255
         Left            =   -71760
         TabIndex        =   233
         Top             =   1200
         Width           =   4815
      End
   End
   Begin Crystal.CrystalReport CrystalReport6 
      Left            =   0
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CR_CD34 
      Left            =   360
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   1320
      TabIndex        =   59
      Top             =   0
      Width           =   9135
   End
   Begin VB.Label LBL_CP 
      AutoSize        =   -1  'True
      Caption         =   "HAEMATO PATHOLOGY  REPORT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   2880
      TabIndex        =   43
      Top             =   480
      Width           =   4755
   End
End
Attribute VB_Name = "hemcbcreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Dim str As String
Dim testorder As String
Dim rptRSQR As rdoQuery
Dim sel_deptqr As rdoQuery
Dim rptRS As rdoResultset
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
Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset
Dim SELTA1012FLAG1 As Boolean
'''FOR COAGULATION PROFILE
Dim INSCP_QR As rdoQuery
Dim INSCP_RS As rdoResultset
Dim SELCPRNG_QR As rdoQuery
Dim SELCPRNG_RS As rdoResultset
Dim SELCPDETVIEW_QR As rdoQuery
Dim SELCPDETVIEW_RS As rdoResultset
Dim DCBCREPA_RS As rdoResultset

Dim SEL_DHEMCPDET_QR As rdoQuery
Dim SEL_DHEMCPDET_RS As rdoResultset

Dim range_n_qr As rdoQuery
Dim range_n_rs As rdoResultset
'Dim SELEMPLIB_QR As rdoQuery
'Dim SELEMPLIB_RS As rdoResultset

Dim I As Integer
Dim UPDRTIMEC_QR As rdoQuery
Dim SELDCBCDET_QR As rdoQuery
Dim SELDCBCDET_RS As rdoResultset

'Dim updhinterface_qr As rdoQuery

'Dim insdfsrep_qr As rdoQuery
'Dim insdgrossdesc_qr As rdoQuery
'Dim insdfs_qr As rdoQuery

Dim UPDDHEMCBCDET_QR As rdoQuery

Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG1 As Boolean
'Dim selempbasinfo1_qr As rdoQuery
'Dim selempbasinfo1_rs As rdoResultset
Dim updrrequisitn_qr As rdoQuery
Dim updrrequisitn1_qr As rdoQuery
Dim INSDADDENDUM_QR As rdoQuery
Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim SELDREQUISITNFLAG1 As Boolean

Dim INSDHEMCBCDET_QR As rdoQuery
Dim HAGE As String
Dim HBSEX As String
Dim HBHV As String
Dim HBLV As String
'Dim RANGE
Dim SELDHEMCBCQR As rdoQuery
Dim SELDHEMCBCRS As rdoResultset
Dim SELDHEMCBCQRFLAG As Boolean
Dim HBFLAG As Boolean
Dim RBCFLAG As Boolean
Dim TLCFLAG As Boolean
Dim PLATELETSFLAG As Boolean
Dim NEUTROPHILFLAG As Boolean
Dim LYMPHOCYTEFLAG As Boolean
Dim MONOCYTESFLAG As Boolean
Dim EOSINOPHILFLAG As Boolean
Dim BASOPHILSFLAG As Boolean
Dim IMMGRANULOFLAG As Boolean
Dim HCTFLAG As Boolean
Dim MCHFLAG As Boolean
Dim MCVFLAG As Boolean
Dim MCHCFLAG As Boolean
Dim MPVFLAG As Boolean
Dim RDWFLAG As Boolean

Dim SELHINTFVAL_QR As rdoQuery
Dim SELHINTFVAL_RS As rdoResultset

Dim SELDHEMCBCDET_QR As rdoQuery
Dim SELDHEMCBCDET_RS As rdoResultset
Dim SELDHEMCBCDETFLAG As Boolean

Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim flagsave As Integer
Dim ADDENDUMFLAG As Boolean
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

Dim selMA101QR As rdoQuery
Dim SELMA101RS As rdoResultset
Dim SELMA101FLAG As Boolean

Dim UPDDREQDRNAME_QR As rdoQuery
Dim UPDDREQDRNAME_RS As rdoResultset
Dim SELDREQDRNAME_QR As rdoQuery
Dim SELDREQDRNAME_RS As rdoResultset

Dim SELDMULPROCQR As rdoQuery
Dim seldmulprocrs As rdoResultset
Dim seldmulprocflag As Boolean

'Dim updrrequisitn1_qr As rdoQuery
Dim SELDREQNOQR As rdoQuery
Dim SELDREQNORS As rdoResultset
Dim SELDREQNOFLAG As Boolean


Dim w_impression As Application     'Dim w_imp As Application    ', w_rhythm As Application, w_stsegment As Application, w_twave As Application,
'Dim d_impression As Document ', d_rhythm As Document, d_stsegment As Document, d_twave As Document, d_impression As Document
Dim d_impression As Object
Dim w_comments As Application
Dim d_comments As Document

Dim w_rbcmorphology As Application
'Dim d_rbcmorphology As Document
Dim d_rbcmorphology As Object

Dim w_pltmorphology As Application
'Dim d_pltmorphology As Document
Dim d_pltmorphology As Object

Dim W_ADDENDUM As Application
'Dim d_ADDENDUM As Document
Dim d_ADDENDUM As Object

Dim INSDBMAREPORT_QR As rdoQuery
Dim UPDDBMAREPORT_QR As rdoQuery

Dim SELDBMAREPORT_QR As rdoQuery
Dim SELDBMAREPORT_RS As rdoResultset
Dim SELDBMAREPORTFLAG As Boolean

Dim INSCBCSUBREP_QR As rdoQuery
Dim SELCBCSUBREP_QR As rdoQuery
Dim SELCBCSUBREP_RS As rdoResultset
Dim SELCBCSUBREPFLAG As Boolean
Dim diffcountflag As Boolean
Dim RNAGE As String
Dim SS As String
Dim CYTOCHEMFLAG As Boolean


''*****prepathno**********
Dim PREPATH_QR As rdoQuery
Dim PREPATH_RS As rdoResultset
Dim PREPATH_FLAG As Boolean
'''******REPORT********
Dim seldrequisitn_qr10 As rdoQuery
Dim seldrequisitn_rs10 As rdoResultset
Dim seldrequisitn_flag10 As Boolean

Dim seldrequisitn_qr11 As rdoQuery
Dim seldrequisitn_rs11 As rdoResultset
Dim seldrequisitn_flag11 As Boolean


Dim X
Dim DEPT1 As String
Dim MODALITY1 As String
Dim REQNO1 As String


Dim SELRREQNOQR As rdoQuery
Dim SELRREQNORS As rdoResultset
Dim SELRREQNOFLAG As Boolean

Dim SELDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_RS As rdoResultset
Dim SELDADDENDUM_FLAG As Boolean

'Dim SELPATHDREQUISITN_QR As rdoQuery
'Dim selpathdrequisitn_rs As rdoResultset

Dim SELDFREPORT_QR As rdoQuery
Dim SELDFREPORT_RS As rdoResultset

'Dim selta1012_qr As rdoQuery
'Dim selta1012_rs As rdoResultset

Dim SELMA301QR As rdoQuery
Dim SELMA301RS As rdoResultset

Dim SELEMPBASINFO_QR As rdoQuery
Dim SELEMPBASINFO_RS As rdoResultset

''HLREPORT
Dim SELDCBCREPORT_QR As rdoQuery
Dim SELDCBCREPORT_RS As rdoResultset
Dim SELDCBCREPORT_FLAG As Boolean
''HLREPORT
Dim SELDBMAREP_QR As rdoQuery
Dim SELDBMAREP_RS As rdoResultset
Dim SELDBMAREP_FLAG As Boolean


'''******REPORT********
Dim seldissecentr11_RS As rdoResultset
Dim seldissecentr11_QR As rdoQuery

''******SECURITY***********

Dim LOGINUSRNAME As String
Dim LOGINTYP As String
Dim LOGINUSRID As String

Dim REPWTRNRS As rdoResultset
Dim REPWTRN_QR As rdoQuery

Dim REPWTRIRS As rdoResultset
Dim CURUSRDET_QR As rdoQuery
Dim CURUSRDET_RS As rdoResultset

Dim SELFINALBY_QR As rdoQuery
Dim SELFINALBY_RS As rdoResultset
Dim SELCPDET_QR As rdoQuery
Dim SELCPDET_RS As rdoResultset
Dim UPDCPDET_QR As rdoQuery
Dim UPDCPDET_RS As rdoResultset

''******SECURITY***********
'' hemato interface
Dim SELDHEMCYTO_QR As rdoQuery
Dim SELDHEMCYTO_RS As rdoResultset
Dim selcinterface_qr As rdoQuery
Dim selhinterface_qr As rdoQuery
Dim updhinterface_qr As rdoQuery
Dim updcinterface_qr As rdoQuery
Dim selcinterface_rs As rdoResultset
Dim selhinterface_rs As rdoResultset
Dim inscbcsm_qr As rdoQuery
Dim selcbcsm_qr As rdoQuery
Dim updcbcsm_qr As rdoQuery
Dim selcbcsm_rs As rdoResultset

''' KANCHAN
Dim SELDCBCREPA_QR As rdoQuery
Dim SELDCBCREPA_RS As rdoResultset
Dim INSDCBCREPA_QR As rdoQuery
Dim UPDDCBCREPA_QR As rdoQuery
Dim SELDCBCFLAG As Boolean
Dim UPDRREQ_QR As rdoQuery
Dim SELDCBC_QR As rdoQuery
Dim SELDCBC_RS As rdoResultset

Dim VEWDHEMCBCDET_QR As rdoQuery
Dim VEWDHEMCBCDET_RS As rdoResultset

Dim CNTFLG As Boolean
'Dim UPDDHEMCBCDET_QR As rdoQuery

Dim UPDTADNDMTIME_QR As rdoQuery 'Shraddha Pandit 30/03/2009

'***** Shraddha Pandit 23/09/09 *****
Dim UPD_HCD34_QR As rdoQuery
Dim sel_rtimelog_qr As rdoQuery
Dim sel_rtimelog_rs As rdoResultset

Dim SEL_DHINTER_QR As rdoQuery '09/03/11
Dim SEL_DHINTER_RS As rdoResultset '09/03/11
Dim sel_dhemcbc_qr As rdoQuery
Dim sel_dhemcbc_rs As rdoResultset

Dim SELSTAFFINFO_QR As rdoQuery '09/03/11
Dim SELSTAFFINFO_RS As rdoResultset '09/03/11

Dim SEL_USER_QR As rdoQuery '27/04/11
Dim SEL_USER_RS As rdoResultset '27/04/11
'************************************

' hemato interface

Public Sub cmb_dept_Click()
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
    Else
    cmb_modality.clear
        If Trim(txt_deptcode.Text) = "FED" Then
        cmb_modality.AddItem "Emergency Heamatopathology" & "$" & "HL"
        cmb_modality.ListIndex = 0
        Else
        cmb_modality.ListIndex = -1
        End If
    End If
    
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

''''''ADDED BY KANCHAN

Public Sub VIEW_CBC()

VEWDHEMCBCDET_QR(0) = req_chglen(cmb_reqno.Text)
Set VEWDHEMCBCDET_RS = VEWDHEMCBCDET_QR.OpenResultset(1, 3)

If VEWDHEMCBCDET_RS.RowCount > 0 Then
    While Not VEWDHEMCBCDET_RS.EOF
        If Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "HAEMOGLOBIN" Then
            TXT_HB.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "RBC COUNT" Then
            TXT_RBC.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
        
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "WBC COUNT" Then
            TXT_TWBC.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "PLATELETS" Then
            txt_platelets.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "NEUTROPHIL %" Then
            txt_neutrophils.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
        
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "LYMPHOCYTE %" Then
            txt_lymphocytes.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "MONOCYTE %" Then
            txt_monocytes.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
        
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "EOSINOPHIL %" Then
            txt_eosinophils.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
        
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "BASOPHIL %" Then
            txt_basophils.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "HAEMATOCRIT" Then
            TXT_HCT.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "MCV" Then
            TXT_MCV.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "MCH" Then
            TXT_MCH.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "MCHC" Then
            TXT_MCHC.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "RDW" Then
            TXT_RDW.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "MPV" Then
            TXT_MPV.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "NEUTROPHIL ABS" Then
            txt_ntphabs.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "LYMPHOCYTE ABS" Then
            txt_lymabs.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "MONOCYTE ABS" Then
            txt_mncyteabs.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "EOSINOPHIL ABS" Then
            txt_esnabs.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
            
        ElseIf Trim(VEWDHEMCBCDET_RS("PRO_CODE")) = "BASOPHIL ABS" Then
            txt_bsphabs.Text = Trim(VEWDHEMCBCDET_RS("OBSERVALUE"))
        End If
        
        VEWDHEMCBCDET_RS.MoveNext
    Wend
End If

End Sub

Private Sub cmb_dept_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_dept) = "" Then
    For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
        If Right(Trim(cmb_dept.List(ind_cbodept)), 3) = "UZZ" Then
            cmb_dept.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub


Private Sub cmb_machname_Change()
 'cmb_machname_Click
End Sub

Private Sub cmb_machname_Click()
'If Trim(cmb_machname.Text) <> "" Then
'Option1.Value = False
'Option2.Value = False
'Option3.Value = False
'End If
End Sub

Public Sub cmb_modality_Click()
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
MODALITY1 = Right(cmb_modality.Text, (l - moda))
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub


Private Sub cmb_modality_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_modality) = "" Then
    For ind_cbodept = 0 To (cmb_modality.ListCount - 1)
        If Trim(cmb_modality.List(ind_cbodept)) = "HAEMATOLOGY$HL" Then
            cmb_modality.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmb_mpo_Click()
    If cmb_mpo.Text = "Negative" Then
        TXT_MPO.Visible = False
    Else
        TXT_MPO.Visible = True
    End If
End Sub

Private Sub cmb_nse_Change()
    If cmb_nse.Text = "Negative" Then
        TXT_NSE.Visible = False
    Else
        TXT_NSE.Visible = True
    End If
End Sub

Private Sub CMB_PREPATHNO_Click()
On Error GoTo Errhandler
Dim X
X = InStr(1, Trim(CMB_PREPATHNO.Text), "$")
REQNO1 = Mid(Trim(CMB_PREPATHNO.Text), 1, X - 1)
DEPT1 = Mid(Trim(REQNO1), 1, 3)
MODALITY1 = Mid(Trim(REQNO1), 5, 2)
If DEPT1 = "FZZ" And MODALITY1 = "SP" Then
    SPREPORT
ElseIf DEPT1 = "FZZ" And MODALITY1 = "FA" Then
    FAREPORT
ElseIf DEPT1 = "FZZ" And (MODALITY1 = "CG" Or MODALITY1 = "CN") Then
    CGCNREPORT
ElseIf DEPT1 = "UZZ" And MODALITY1 = "BM" Then
    BONEMARROW
ElseIf DEPT1 = "UZZ" And MODALITY1 = "BF" Then
    BFREPORT
ElseIf DEPT1 = "UZZ" And MODALITY1 = "HL" Then
    HLREPORT
End If
Exit Sub
Errhandler:
MsgBox Err.Description, vbCritical, "DIS"
Exit Sub
End Sub

Public Sub cmb_reqno_Click()
CMB_range.Text = "select"
Dim STR6, STR7
Dim range_n_qr3 As rdoQuery
Dim range_n_rs3 As rdoResultset
STR6 = "select OBSERTYPE from dislib.DHEMCpdet where reqno='" & req_chglen(cmb_reqno.Text) & "' and PROCCODE='D-DIMER'"
       Set range_n_qr3 = rdb.CreateQuery("", STR6)

Set range_n_rs3 = range_n_qr3.OpenResultset(1, 1)

If range_n_rs3.RowCount > 0 Then
 'While range_n_rs.Read
 CMB_range.Text = range_n_rs3(0)
 'range_n_rs.MoveNext
 'Wend
End If

Dim prpathno As String
Dim ADDENDUM As String
Dim entr, CytoFlag As String
'On Error GoTo MSG:
'' setting of tabs based on dhemcyto1 table
'      SSTab1.TabEnabled(0) = False
'      SSTab1.TabEnabled(1) = False
'      SSTab1.TabEnabled(2) = False
'      SSTab1.TabEnabled(3) = False
'      SSTab1.TabEnabled(4) = False
      
'************ Shraddha Pandit 30/03/09 ****************
txt_addendum.Enabled = False
cmb_addendumby.Enabled = False
'******************************************************
      
      'FOR CHEKING CODES
        If seldmulprocflag = False Then
            SELDMULPROCQR(0) = req_chglen(cmb_reqno.Text)
            Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
            seldmulprocflag = True
        Else
            SELDMULPROCQR(0) = req_chglen(cmb_reqno.Text)
            seldmulprocrs.MoreResults
            seldmulprocrs.Requery
        End If
        If seldmulprocrs.RowCount > 0 Then
            seldmulprocrs.MoveFirst
            While Not seldmulprocrs.EOF
                If seldmulprocrs("pro_code") = "U752" Then
                   CytoFlag = "YES"
                   SSTab1.TabEnabled(3) = True
                End If
                seldmulprocrs.MoveNext
            Wend
       End If
       
       SELDHEMCYTO_QR(0) = req_chglen(cmb_reqno.Text)
       Set SELDHEMCYTO_RS = SELDHEMCYTO_QR.OpenResultset(1, 1)
      If SELDHEMCYTO_RS.RowCount > 0 Then
           If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("HEM")) = "YES" Then
             SSTab1.TabEnabled(0) = True
             SSTab1.TabEnabled(1) = True
             SSTab1.TabEnabled(3) = True
           End If
           
           '***** Shraddha 06/05/2011 *****
'           If Trim(SELDHEMCYTO_RS("MPO")) = "YES" Then
'            SSTab1.TabEnabled(1) = True
'           End If
'           '*******************************
'           If Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
'             SSTab1.TabEnabled(2) = True
'           End If
'           If Trim(SELDHEMCYTO_RS("CYTO")) = "YES" Then
'             SSTab1.TabEnabled(3) = True
'           End If
'           If Trim(SELDHEMCYTO_RS("SMAR")) = "YES" Then
'             SSTab1.TabEnabled(4) = True
'           End If
           
      End If
 '' seting of tabs based on dhemcyto1 table over
If Opt_patient.Value = True Then
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
            If Len(Trim(SELTA101_RS("BIRTH_DATE"))) <> 0 Then
                    'PATYPE = Trim(SELTA101_RS("PAT_TYPE"))
                   YEAR1 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), DATE11)
                   
                   Dim n As Integer
                   n = YEAR1 / 12
                   txt_age.Text = n
               End If
               
               If Trim(txt_age.Text) = 0 Then
                  Dim age As String
                  age = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                  age = "." & (age Mod 12)
                     If age = ".0" Then
                        age = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                        age = ".0" & (age) ' Mod 30)
                     End If
                  txt_age.Text = Trim(age)
               End If
            'txt_age = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
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
                'lbl_bed.Visible = True
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_status.Text = "OP"
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
               ' lbl_bed.Visible = False
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
                TXT_PATH.Text = "" & Trim(selpathdrequisitn_rs("PATHNO"))
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
                txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
                       
                ENABL ("ENA")
                
''        If seldmulprocflag = False Then
''            SELDMULPROCQR(0) = req_chglen(cmb_reqno.Text)
''            Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
''            seldmulprocflag = True
''        Else
''            SELDMULPROCQR(0) = req_chglen(cmb_reqno.Text)
''            seldmulprocrs.MoreResults
''            seldmulprocrs.Requery
''        End If
''        If seldmulprocrs.RowCount > 0 Then
''            While Not seldmulprocrs.EOF
''                If seldmulprocrs("pro_code") = "F828" Then
''                    CYTOCHEMFLAG = False
''                Else
''                    CYTOCHEMFLAG = True
''                End If
''                seldmulprocrs.MoveNext
''            Wend
''        End If
        
'        If seldissecentrflag1 = False Then
'            Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
'            seldissecentrflag1 = True
'        Else
'            seldissecentr1_rs.MoreResults
'            seldissecentr1_rs.Requery
'        End If
'        If seldissecentr1_rs.RowCount > 0 Then
'
'            entr = cmb_enterby.Text
'            cmb_enterby.clear
'            While Not seldissecentr1_rs.EOF
'                cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
'                seldissecentr1_rs.MoveNext
'            Wend
'            seldissecentr1_rs.MovePrevious
'            cmb_enterby.Text = seldissecentr1_rs("CCNAME")
'
'        End If

'        If seldissecfinflag = False Then
'            Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
'            seldissecfinflag1 = True
'        Else
'            seldissecfin1_rs.MoreResults
'            seldissecfin1_rs.Requery
'        End If
'        If seldissecfin1_rs.RowCount > 0 Then
'            entr = ""
'            entr = cmb_finalby.Text
'            cmb_finalby.clear
'            While Not seldissecfin1_rs.EOF
'                cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
'                seldissecfin1_rs.MoveNext
'            Wend
'            cmb_finalby.Text = entr
'            entr = ""
'        End If
            
'' hemato enabling of tabs and interface fetching
'      SSTab1.TabEnabled(0) = False
'      SSTab1.TabEnabled(1) = False
'      SSTab1.TabEnabled(2) = False
'      SSTab1.TabEnabled(3) = False
'      SSTab1.TabEnabled(4) = False
'      SELDHEMCYTO_QR(0) = req_chglen(cmb_reqno.Text)
'       Set SELDHEMCYTO_RS = SELDHEMCYTO_QR.OpenResultset(1, 1)
'      If SELDHEMCYTO_RS.RowCount > 0 Then
'           If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("HEM")) = "YES" Then
'             SSTab1.TabEnabled(0) = True
'             SSTab1.TabEnabled(1) = True
'           End If
'           If Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
'             SSTab1.TabEnabled(2) = True
'           End If
'           If Trim(SELDHEMCYTO_RS("CYTO")) = "YES" Then
'             SSTab1.TabEnabled(3) = True
'           End If
'           If Trim(SELDHEMCYTO_RS("SMAR")) = "YES" Then
'             SSTab1.TabEnabled(4) = True
'           End If
'
'      End If
'' hemato  interface fetching
        selhinterface_qr(0) = Mid(req_chglen(cmb_reqno.Text), 6, 13)
        Set selhinterface_rs = selhinterface_qr.OpenResultset(1, 1)
        If selhinterface_rs.RowCount > 0 Then
         TXT_HB = Val(selhinterface_rs("hb"))
         TXT_RBC = Val(selhinterface_rs("rbc"))
         TXT_TWBC = Val(selhinterface_rs("tlc"))
         txt_platelets = Val(selhinterface_rs("platelets"))
         txt_immaturegranulocytes = Val(selhinterface_rs("immgranulo"))
         TXT_HCT = Val(selhinterface_rs("hct"))
         TXT_MCH = Val(selhinterface_rs("mch"))
         TXT_MCHC = Val(selhinterface_rs("mchc"))
         TXT_MCV = Val(selhinterface_rs("mcv"))
         TXT_MPV = Val(selhinterface_rs("mpv"))
         TXT_RDW = Val(selhinterface_rs("rdw"))
         txt_neutrophils = Val(selhinterface_rs("neutrophil"))
         txt_lymphocytes = Val(selhinterface_rs("lymphocyte"))
         txt_monocytes = Val(selhinterface_rs("monocytes"))
         txt_eosinophils = Val(selhinterface_rs("eosinophil"))
         txt_basophils = Val(selhinterface_rs("basophils"))
         txt_ntphabs.Text = (TXT_TWBC.Text) * ((txt_neutrophils.Text) / 100)
         txt_lymabs.Text = (TXT_TWBC.Text) * ((txt_lymphocytes.Text) / 100)
         txt_mncyteabs.Text = (TXT_TWBC.Text) * ((txt_monocytes.Text) / 100)
         txt_esnabs.Text = (TXT_TWBC.Text) * ((txt_eosinophils.Text) / 100)
         txt_bsphabs.Text = (TXT_TWBC.Text) * ((txt_basophils.Text) / 100)
        End If
        selcinterface_qr(0) = req_chglen(cmb_reqno.Text)
        Set selcinterface_rs = selcinterface_qr.OpenResultset(1, 1)
        Do While selcinterface_rs.EOF = False
          If selcinterface_rs("testcode") = "U708" Then
          TXT_PT(0).Text = selcinterface_rs("TESTRESULT")
           
          ElseIf selcinterface_rs("testcode") = "U710" Then
            TXT_PTTK = selcinterface_rs("TESTRESULT")
          ElseIf selcinterface_rs("testcode") = "U714" Then
            TXT_FDPDMIER = selcinterface_rs("TESTRESULT")
            
          ElseIf selcinterface_rs("testcode") = "U715" Then
            TXT_FIBRINOGEN = selcinterface_rs("TESTRESULT")
          End If
          
          selcinterface_rs.MoveNext
          'selcinterface_rs.MoveNext
          
        Loop
        '''*********** VIEW COAGULATION PROFILE
       
             
        
        
        
        
        
        
'' hemato  interface fetching ends

' ends
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
            If Len(Trim(SELTA101_RS("BIRTH_DATE"))) <> 0 Then
                    'PATYPE = Trim(SELTA101_RS("PAT_TYPE"))
                   YEAR1 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), DATE11)
                   
                   'Dim N As Integer
                   n = YEAR1 / 12
                   txt_age.Text = n
               End If
               
               If Trim(txt_age.Text) = 0 Then
                  'Dim age As String
                  age = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                  age = "." & (age Mod 12)
                     If age = ".0" Then
                        age = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                        age = ".0" & (age) ' Mod 30)
                     End If
                  txt_age.Text = Trim(age)
               End If
            'txt_age = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), DATE11)
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
                    'MsgBox "Report is commited , You can not Modify"
                    cmd_modify.Enabled = False
                    
                    '*************** Shraddha Pandit 30/03/09 ***************
                    MsgBox "Final Report is Generated, You can't Modify the Final Report ,You can Add ADDENDUMS Only (for Surface Marker)"
                    ADDENDUMFLAG = True
                    'cmd_modify.Enabled = True
                    cmd_commit.Enabled = False
                    
                    '***** Shraddha 31/10/2011 *****
                    txt_addendum.Enabled = True
                    cmb_addendumby.Enabled = True
                    cmd_save.Enabled = True
                    cmd_commit.Enabled = False
                    '********************************************************
                                       
                    cmd_mdc.Enabled = True
                End If
                cmb_enterby.Text = "" & REFDR_RS("ENTERBY")
                CMB_PROVISIONAL.Text = "" & REFDR_RS("PREPORTBY")
                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    cmb_finalby.Text = ""
                Else
                    cmb_finalby.Text = REFDR_RS("FREPORTBY")
                End If
                If IsNull(Trim(REFDR_RS("PREPORTBY"))) Then
                    regcbo.Text = ""
                Else
                    regcbo.Text = REFDR_RS("PREPORTBY")
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
                'lbl_bed.Visible = False
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                'lbl_bed.Visible = False
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
                TXT_PATH.Text = "" & Trim(selpathdrequisitn_rs("PATHNO"))
                If selpathdrequisitn_rs.EOF = False Then
                    'txt_relclnhis.Visible = True
                    
                    'txt_relclnhis.Text = " " & selpathdrequisitn_rs("NMETERIALS")
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
            
'            If seldmulprocflag = False Then
'                SELDMULPROCQR(0) = req_chglen(cmb_reqno.Text)
'                Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
'                seldmulprocflag = True
'            Else
'                SELDMULPROCQR(0) = req_chglen(cmb_reqno.Text)
'                seldmulprocrs.MoreResults
'                seldmulprocrs.Requery
'            End If
'        If seldmulprocrs.RowCount > 0 Then
'            While Not seldmulprocrs.EOF
'
'                If seldmulprocrs("pro_code") = "F828" Then
'                    CYTOCHEMFLAG = False
'                Else
'                    CYTOCHEMFLAG = True
'                End If
'                seldmulprocrs.MoveNext
'            Wend
'        End If
            
            
            If SELDBMAREPORTFLAG = False Then
                SELDBMAREPORT_QR(0) = req_chglen(cmb_reqno.Text)
                Set SELDBMAREPORT_RS = SELDBMAREPORT_QR.OpenResultset(1, 3)
                SELDBMAREPORTFLAG = True
            Else
                SELDBMAREPORT_QR(0) = req_chglen(cmb_reqno.Text)
                SELDBMAREPORT_RS.MoreResults
                SELDBMAREPORT_RS.Requery
            End If
            ENABL ("DIS")
            VIEW_CBC
            If SELDBMAREPORT_RS.RowCount > 0 Then
                    '**** cbc counter
                    If SELCBCSUBREPFLAG = False Then
                        SELCBCSUBREP_QR(0) = req_chglen(cmb_reqno.Text)
                        Set SELCBCSUBREP_RS = SELCBCSUBREP_QR.OpenResultset(1, 3)
                        SELCBCSUBREPFLAG = False
                    Else
                        SELCBCSUBREP_QR(0) = req_chglen(cmb_reqno.Text)
                        SELCBCSUBREP_RS.MoreResults
                        SELCBCSUBREP_RS.Requery
                    End If
                    If SELCBCSUBREP_RS.RowCount > 0 Then
                        SELCBCSUBREP_RS.MoveLast
                        TXT_HB.Text = " " & SELCBCSUBREP_RS("hb")
                        TXT_RBC.Text = "" & SELCBCSUBREP_RS("rbc")
                        TXT_TWBC.Text = " " & SELCBCSUBREP_RS("tlc")
                        txt_platelets.Text = " " & SELCBCSUBREP_RS("platelets")
                    End If
           '''  surface marker data fetching for view
            selcbcsm_qr(0) = req_chglen(cmb_reqno.Text)
               Set selcbcsm_rs = selcbcsm_qr.OpenResultset(1, 1)
                If selcbcsm_rs.RowCount > 0 Then
                    SURFACEFORM.TXT_ANTIMPO = selcbcsm_rs("antimpo") '
                    SURFACEFORM.TXT_CD10 = selcbcsm_rs("cd10") '" "
                    SURFACEFORM.txt_cd103 = selcbcsm_rs("cd103") '  " "
                    SURFACEFORM.TXT_CD117 = selcbcsm_rs("cd117") ' " "
                    SURFACEFORM.TXT_CD11C = selcbcsm_rs("cd11c")  '" "
                    
                    SURFACEFORM.txt_cd13 = selcbcsm_rs("cd13")  '" "
                    SURFACEFORM.TXT_CD138 = selcbcsm_rs("cd138")  '" "
                    SURFACEFORM.txt_cd14 = selcbcsm_rs("cd14")  '" "
                    SURFACEFORM.txt_CD15 = selcbcsm_rs("cd15")  '" "
                    SURFACEFORM.TXT_CD19 = selcbcsm_rs("cd19")  ' " "
                    SURFACEFORM.TXT_CD2 = selcbcsm_rs("cd2")   '" "
                    SURFACEFORM.TXT_CD20 = selcbcsm_rs("cd20")  ' ""
                    SURFACEFORM.TXT_CD22 = selcbcsm_rs("cd22")  ' ""
                    SURFACEFORM.TXT_CD23 = selcbcsm_rs("cd23")  ' ""
                    SURFACEFORM.TXT_CD25 = selcbcsm_rs("cd25")  '""
                    SURFACEFORM.TXT_CD3 = selcbcsm_rs("cd3")  ' " "
                    SURFACEFORM.txt_CD33 = selcbcsm_rs("cd33")  '""
                    SURFACEFORM.TXT_CD34 = selcbcsm_rs("cd34")  '" "
                    SURFACEFORM.txt_cd36 = selcbcsm_rs("cd16")  '" "
                    SURFACEFORM.TXT_CD38 = selcbcsm_rs("cd38") ' " "
                    SURFACEFORM.TXT_CD4 = selcbcsm_rs("cd4")  ' " "
                    SURFACEFORM.txt_cd41 = selcbcsm_rs("cd41")  ' " "
                    SURFACEFORM.TXT_CD42 = selcbcsm_rs("cd42")  ' " "
                    SURFACEFORM.TXT_CD5 = selcbcsm_rs("cd5")  ' " "
                    SURFACEFORM.TXT_CD56 = selcbcsm_rs("cd56")  '" "
                    SURFACEFORM.TXT_CD61 = selcbcsm_rs("cd61")  '" "
                    SURFACEFORM.TXT_CD68 = selcbcsm_rs("cytocd68")  '" "
                    SURFACEFORM.TXT_CD7 = selcbcsm_rs("cd7")   '" "
                    SURFACEFORM.txt_CD79B = selcbcsm_rs("cd79b")  '" "
                    SURFACEFORM.TXT_CD8 = selcbcsm_rs("cd8")  '""
                    SURFACEFORM.txt_cytocd221(1) = selcbcsm_rs("cytocd22")  '" "
                    SURFACEFORM.TXT_CYTOCD3 = selcbcsm_rs("cytocd3")  ' " "
                    SURFACEFORM.txt_cytocd79a = selcbcsm_rs("cytocd79a")  '" "
                    SURFACEFORM.TXT_CYTOMU = selcbcsm_rs("cytomu")  ' " "
                    SURFACEFORM.TXT_DBA44 = selcbcsm_rs("dba44")  '" "
                    SURFACEFORM.TXT_FMC7 = selcbcsm_rs("fmc7")  '" "
                    SURFACEFORM.txt_Glycophorina = selcbcsm_rs("glycophoin")  '" "
                    SURFACEFORM.TXT_HLADR = selcbcsm_rs("hladr")  '" "
                    SURFACEFORM.txt_igm = selcbcsm_rs("igm")  '" "
                    SURFACEFORM.txt_kappa = selcbcsm_rs("kappa")  '" "
                    SURFACEFORM.txt_lanbda = selcbcsm_rs("lamda")  '" "
                    SURFACEFORM.txt_smig = selcbcsm_rs("smig")  '" "
                    SURFACEFORM.TXT_TDT = selcbcsm_rs("tdt")  ' " "
                    'added by milind on 07-08-2006
                    SURFACEFORM.txt_clnd1 = selcbcsm_rs("clnd1")
                    SURFACEFORM.txt_tcrab = selcbcsm_rs("tcrab")
                    SURFACEFORM.txt_tcrgd = selcbcsm_rs("tcrgd")
                    SURFACEFORM.TXT_SMPCTG = selcbcsm_rs("smpctg")
                    
                        If selcbcsm_rs("optlsp") = "Y" Then
                            SURFACEFORM.opt_lsp.Value = True
                        ElseIf selcbcsm_rs("optlsp") = "N" Then
                            SURFACEFORM.opt_lsp.Value = False
                        End If
                    
                        If selcbcsm_rs("optscp") = "Y" Then
                            SURFACEFORM.opt_ssp.Value = True
                        ElseIf selcbcsm_rs("optscp") = "N" Then
                            SURFACEFORM.opt_ssp.Value = False
                        End If
                        
                        
                    'upto here
                 End If
                          
                 SELDREQDRNAME_QR(0) = req_chglen(cmb_reqno.Text)
                 Set SELDREQDRNAME_RS = SELDREQDRNAME_QR.OpenResultset(1, 3)
                 If SELDREQDRNAME_RS.RowCount > 0 Then
                        If Trim(SELDREQDRNAME_RS("LATERALITS")) = "Dr. Ashok Kumar" Then
                            SURFACEFORM.opt_drak.Value = True
                            SURFACEFORM.opt_mryb.Value = False
                            SURFACEFORM.opt_sitaram.Value = False
                        ElseIf Trim(SELDREQDRNAME_RS("METASTASIS")) = "Mr.Y. Badrinath" Then
                            SURFACEFORM.opt_drak = False
                            SURFACEFORM.opt_mryb.Value = True
                            SURFACEFORM.opt_sitaram.Value = False
                        ElseIf Trim(SELDREQDRNAME_RS("METASTASIS")) = "Mr. Sitaram" Then
                            SURFACEFORM.opt_drak = False
                            SURFACEFORM.opt_mryb.Value = False
                            SURFACEFORM.opt_sitaram.Value = True
                        End If
                 End If
                 ''******* VIEW COAGULATION PROFILE***********
                 SELCPDETVIEW_QR(0) = req_chglen(cmb_reqno.Text)
                 Set SELCPDETVIEW_RS = SELCPDETVIEW_QR.OpenResultset(2, 1)
                 If Not SELCPDETVIEW_RS.EOF Then
                        Do While Not SELCPDETVIEW_RS.EOF '1
                        If Trim(SELCPDETVIEW_RS("PROCCODE")) = "PT" Then
                            TXT_PT(0).Text = Trim(SELCPDETVIEW_RS("OBSERVALUE"))
                            txt_ptctl.Text = Trim(SELCPDETVIEW_RS("CONTROL"))
                        
                        End If
                        If Trim(SELCPDETVIEW_RS("PROCCODE")) = "INR" Then
                            txt_inr.Text = Trim(SELCPDETVIEW_RS("OBSERVALUE"))
                            
                        
                        End If
                        If Trim(SELCPDETVIEW_RS("PROCCODE")) = "APTT" Then
                            TXT_PTTK.Text = Trim(SELCPDETVIEW_RS("OBSERVALUE"))
                            txt_apttctl.Text = Trim(SELCPDETVIEW_RS("CONTROL"))
                        
                        End If
                        If Trim(SELCPDETVIEW_RS("PROCCODE")) = "TT" Then
                            TXT_TT.Text = Trim(SELCPDETVIEW_RS("OBSERVALUE"))
                            txt_ttctl.Text = Trim(SELCPDETVIEW_RS("CONTROL"))
                        
                        End If
                        If Trim(SELCPDETVIEW_RS("PROCCODE")) = "FIBRINOGEN" Then
                            TXT_FIBRINOGEN.Text = Trim(SELCPDETVIEW_RS("OBSERVALUE"))
                        
                        End If
                        If Trim(SELCPDETVIEW_RS("PROCCODE")) = "D-DIMER" Then
                            TXT_FDPDMIER.Text = Trim(SELCPDETVIEW_RS("OBSERVALUE"))
                        
                        End If
'                         If Trim(SELCPDETVIEW_RS("MACHINE")) = "ACL 100" Then
''                            Option1.Value = True
''                             Option2.Value = False
''                             Option3.Value = False
'                        ElseIf Trim(SELCPDETVIEW_RS("MACHINE")) = "ACL 200" Then
''                            Option1.Value = False
''                            Option3.Value = False
''                            Option2.Value = True
'                           ElseIf Trim(SELCPDETVIEW_RS("MACHINE")) = "ELP01" Then
'                            Option1.Value = False
'                             Option2.Value = False
'                            Option3.Value = True
'                            Else
                        If IsNull(SELCPDETVIEW_RS("MACHINE")) = False Then
                            cmb_machname.Text = Trim(SELCPDETVIEW_RS("MACHINE"))
                            End If
'                        End If
                        SELCPDETVIEW_RS.MoveNext
                    Loop
                 End If
                 
                 
                 
                 
                 
                 
                 
                 
           '''  surface marker data fetching for view ends
           
           ''*******ADDENDUM*********************************
            'If SELDADDENDUMFLAG = False Then
             '   SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
              '  Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
               ' SELDADDENDUMFLAG = True
            'Else
            '' HERE
                SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
                SELDADDENDUM_RS.MoreResults
                SELDADDENDUM_RS.Requery
            'End If
            If SELDADDENDUM_RS.RowCount > 0 Then
                ADDENDUM = ""
                While Not SELDADDENDUM_RS.EOF
                    ADDENDUM = ADDENDUM & " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
                    'cmb_addendumby.AddItem SELDADDENDUM_RS("USER_ID")
                    cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
                SELDADDENDUM_RS.MoveNext
                Wend
                txt_addendum.Text = Trim(ADDENDUM)
                ADDENDUMFALG = True
            End If
            ''*******ADDENDUM*********************************

            
                    
                    'txt_neutrophils.Text = " " & SELDBMAREPORT_RS("neutrophil")
                    'txt_lymphocytes.Text = " " & SELDBMAREPORT_RS("lymphocyte")
                    'txt_monocytes.Text = " " & SELDBMAREPORT_RS("monocytes")
                    'txt_eosinophils.Text = " " & SELDBMAREPORT_RS("eosinophil")
                    'txt_basophils.Text = " " & SELDBMAREPORT_RS("basophils")
                    'txt_ntphabs.Text = (TXT_TWBC.Text) * Val((txt_neutrophils.Text) / 100)
                    'txt_lymabs.Text = (TXT_TWBC.Text) * ((txt_lymphocytes.Text) / 100)
                    'txt_mncyteabs.Text = (TXT_TWBC.Text) * ((txt_monocytes.Text) / 100)
                    'txt_esnabs.Text = (TXT_TWBC.Text) * ((txt_eosinophils.Text) / 100)
                    'txt_bsphabs.Text = (TXT_TWBC.Text) * ((txt_basophils.Text) / 100)
                   ' txt_immaturegranulocytes.Text = " " & SELDBMAREPORT_RS("immgranulo")
                    TXT_HCT.Text = " " & SELDBMAREPORT_RS("hct")
                    TXT_MCH.Text = " " & SELDBMAREPORT_RS("mch")
                    TXT_MCV.Text = " " & SELDBMAREPORT_RS("mcv")
                    
                    TXT_MCHC.Text = " " & SELDBMAREPORT_RS("mchc")
                    TXT_MPV.Text = " " & SELDBMAREPORT_RS("mpv")
                    TXT_RDW.Text = " " & SELDBMAREPORT_RS("rdw")
                    TXT_ABNORMAL.Text = " " & SELDBMAREPORT_RS("abnormal")
                    '*** cbc counter
                    '*** cbc manual
                    txt_mneutrophils.Text = " " & SELDBMAREPORT_RS("neutrophi1")
                    txt_mlymphocytes.Text = " " & SELDBMAREPORT_RS("lymphocyt1")
                    txt_mmonocytes.Text = " " & SELDBMAREPORT_RS("monocyte1")
                    txt_meosinophils.Text = " " & SELDBMAREPORT_RS("eosinophi1")
                    txt_mbasophills.Text = " " & SELDBMAREPORT_RS("basophill1")
                    txt_blasts.Text = " " & SELDBMAREPORT_RS("blasts")
                    TXT_PROMYELOCYTESWBC.Text = " " & SELDBMAREPORT_RS("promyelocy")
                    txt_myelocytes.Text = " " & SELDBMAREPORT_RS("myelocytes")
                    txt_metamyelocytes.Text = " " & SELDBMAREPORT_RS("metamyeloc")
                    txt_plasmacells.Text = " " & SELDBMAREPORT_RS("plasmacell")
                    txt_alymphocytes.Text = " " & SELDBMAREPORT_RS("atlymphee")
                    txt_prolymphocytes.Text = " " & SELDBMAREPORT_RS("prolymphoc")
                   ' txt_prolympho.Text = " " & SELDBMAREPORT_RS("PROMYELOCY") 'Shraddha 06/05/2011
                     txt_prolympho.Text = " " & SELDBMAREPORT_RS("prolymphoc") 'SHWETA 30/07/2013
                    TXT_LP.Text = " " & SELDBMAREPORT_RS("fld1") 'Shraddha Pandit 17/10/08
                   ' txt_rc.Text = "" '20/11/08
                    'txt_othertype.Text = " " & DCBCREPA_RS("fld2") 'Shraddha Pandit 21/10/08
                    'txt_malarialparasite.Text = " " & DCBCREPA_RS("maleria") 'Shraddha Pandit 21/10/08
                    txt_Sezarycells.Text = " " & SELDBMAREPORT_RS("sezarycell")
                    txt_hairycells.Text = " " & SELDBMAREPORT_RS("hairycells")
                    txt_othercells.Text = " " & SELDBMAREPORT_RS("othercells")
                    TXT_NORMOBLASTS.Text = " " & SELDBMAREPORT_RS("normboblas")
                    TXT_OTHERCELLSRBC.Text = " " & SELDBMAREPORT_RS("othercellr")
                    txt_rbcmorphology.Text = " " & SELDBMAREPORT_RS("rbcmorpho")
                    Text1.Text = " " & SELDBMAREPORT_RS("pltmorpho")
                    txt_reticcount.Text = " " & SELDBMAREPORT_RS("reticcount")
                    txt_malarialparasite.Text = " " & SELDBMAREPORT_RS("mparasite")
                    txt_esr.Text = " " & SELDBMAREPORT_RS("esr") '18/11/08
                    
                    'TXT_MORPHOLOGY.Text = " " & SELDBMAREPORT_RS("pmorpholog")
                    CMB_MORPHO.Text = " " & SELDBMAREPORT_RS("pmorpholog") 'Shraddha Pandit 13/11/08
                    
                    'TXT_ADEQUACY.Text = " " & SELDBMAREPORT_RS("fld1")
                    VIEWDCBCREPA
                    '*** cbc manual
                    '*** coagulation profile
                    'txt_inr.Text = " " & SELDBMAREPORT_RS("bt")
                    'TXT_PT(0).Text = " " & SELDBMAREPORT_RS("pt")
                    'txt_ptctl.Text = " " & SELDBMAREPORT_RS("platelet1")
                    'TXT_PTTK.Text = " " & SELDBMAREPORT_RS("pttk")
                    'txt_apttctl.Text = " " & SELDBMAREPORT_RS("apttctl")
                    'TXT_FIBRINOGEN.Text = " " & SELDBMAREPORT_RS("fiBrinogen")
                    'TXT_FDPDMIER.Text = " " & SELDBMAREPORT_RS("fdpddimer")
                    'TXT_OTHERS.Text = " " & SELDBMAREPORT_RS("others")
                    'TXT_TT.Text = " " & SELDBMAREPORT_RS("tt")
                    'txt_ttctl.Text = " " & SELDBMAREPORT_RS("ttctl")
                    
                    '*** cogulation profile
                    '*** cytochemistry
                    txt_lapscore.Text = " " & SELDBMAREPORT_RS("lap")
                    txt_lapctl.Text = " " & SELDBMAREPORT_RS("lapctl")
                    
'                    If Trim(SELDBMAREPORT_RS("mpoval")) = "Negative" Then
'                        cmb_mpo.Text = " " & SELDBMAREPORT_RS("mpoval")
'                        TXT_MPO.Visible = False
'                    Else
'                        cmb_mpo.Text = " " & SELDBMAREPORT_RS("mpoval")
'                        TXT_MPO.Visible = True
'                        TXT_MPO.Text = " " & SELDBMAREPORT_RS("mpo")
'                    End If
                    
                    If Trim(SELDBMAREPORT_RS("mpoval")) = "Negative" Then
                        cmbmpo.Text = " " & SELDBMAREPORT_RS("mpoval")
                        'TXT_MPO.Visible = False
                    Else
                        cmbmpo.Text = " " & SELDBMAREPORT_RS("mpoval")
                        'TXT_MPO.Visible = True
                        'TXT_MPO.Text = " " & SELDBMAREPORT_RS("mpo")
                    End If
                    
                    If Trim(SELDBMAREPORT_RS("sbb")) = "Negative" Then
                        cmb_nse.Text = " " & SELDBMAREPORT_RS("sbb")
                        TXT_NSE.Visible = False
                    Else
                        cmb_nse.Text = " " & SELDBMAREPORT_RS("sbb")
                        TXT_NSE.Visible = True
                        TXT_NSE.Text = " " & SELDBMAREPORT_RS("nse")
                    End If
                    
                    TXT_PAS.Text = " " & SELDBMAREPORT_RS("pSA")
                    txt_promonocyte.Text = " " & SELDBMAREPORT_RS("pSA")
                    TXT_TRAP.Text = " " & SELDBMAREPORT_RS("traD")
                    
                    'cytochemistry
                    
                    txt_impression.Text = " " & SELDBMAREPORT_RS("IMPRESSION")
                    Set rs_PBSREPT = rdb.OpenResultset("SELECT * FROM DISLIB.DMULPROCE WHERE REQNO='" & SELDBMAREPORT_RS("reqno") & "' AND (PRO_CODE='U753' OR PRO_CODE='U754') ", 1, 1)
                    If Not rs_PBSREPT.EOF Then
                    txt_1.Visible = True
                    txt_1.Enabled = False
                    txt_1.Text = " " & SELDBMAREPORT_RS("COMMENTS").Value
                    Else
                    txt_comments.Text = " " & SELDBMAREPORT_RS("COMMENTS").Value
                    End If
                    
            End If
            
            If SELDHEMCBCDETFLAG = False Then
                SELDHEMCBCDET_QR(0) = req_chglen(cmb_reqno.Text)
                Set SELDHEMCBCDET_RS = SELDHEMCBCDET_QR.OpenResultset(1, 3)
                SELDHEMCBCDETFLAG = True
            Else
                SELDHEMCBCDET_QR(0) = req_chglen(cmb_reqno.Text)
                SELDHEMCBCDET_RS.MoreResults
                SELDHEMCBCDET_RS.Requery
            End If
            
'////////// COMMENTED BY KANCHAN
            
'            If SELDHEMCBCDET_RS.RowCount > 0 Then
'
'
'                                TXT_HB.Text = " "
'                                TXT_HB.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                TXT_RBC.Text = " "
'                                TXT_RBC.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                TXT_TWBC.Text = " "
'                                TXT_TWBC.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                txt_platelets.Text = " "
'                                txt_platelets.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                txt_neutrophils.Text = " "
'                                txt_neutrophils.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                txt_lymphocytes.Text = " "
'                                txt_lymphocytes.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                txt_monocytes.Text = " "
'                                txt_monocytes.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                txt_eosinophils.Text = " "
'                                txt_eosinophils.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                txt_basophils.Text = " "
'                                txt_basophils.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                txt_ntphabs.Text = " "
'                                txt_ntphabs.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                txt_lymabs.Text = " "
'                                txt_lymabs.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                txt_mncyteabs.Text = " "
'                                txt_mncyteabs.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                txt_esnabs.Text = " "
'                                txt_esnabs.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                txt_bsphabs.Text = " "
'                                txt_bsphabs.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                TXT_HCT.Text = " "
'                                TXT_HCT.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                TXT_MCH.Text = " "
'                                TXT_MCH.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                TXT_MCV.Text = " "
'                                TXT_MCV.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                TXT_MCHC.Text = " "
'                                TXT_MCHC.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                TXT_MPV.Text = " "
'                                TXT_MPV.Text = SELDHEMCBCDET_RS(2)
'
'                                SELDHEMCBCDET_RS.MoveNext
'
'                                TXT_RDW.Text = " "
'                                TXT_RDW.Text = SELDHEMCBCDET_RS(2)
'
'            End If
'''//////////////// UPTO HERE KANCHAN

            VIEW_CBC
End If
End If
If opt_staff.Value = True Then
    STAFF
 
End If

Exit Sub
MSG:
MsgBox Err.Description
End Sub




Public Sub btnenabl(cap As String)
If cap = "NEW" Then
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("SAVE") = "YES" Then
            cmd_save.Enabled = True
        End If
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
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("MODIFY") = "YES" Then
            cmd_modify.Enabled = True
        Else
            cmd_modify.Enabled = False
        End If
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
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("NEW") = "YES" Then
            CMD_NEW.Enabled = True
        Else
            CMD_NEW.Enabled = False
        End If
    End If
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("COMMIT") = "YES" Then
            cmd_commit.Enabled = True
        Else
            cmd_commit.Enabled = False
        End If
    End If
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("VIEW") = "YES" Then
            cmd_view.Enabled = True
        Else
            cmd_view.Enabled = False
        End If
    End If
    'cmd_fullreport.Enabled = False
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("NEW") = "YES" Then
            CMD_NEW.Enabled = True
        Else
            CMD_NEW.Enabled = False
        End If
    End If
    cmd_save.Enabled = False
    cmd_print.Enabled = False
    cmd_commit.Enabled = False
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("VIEW") = "YES" Then
            cmd_view.Enabled = True
        Else
            cmd_view.Enabled = False
        End If
    End If
    'cmd_fullreport.Enabled = False
ElseIf cap = "COMMIT" Then
    cmd_save.Enabled = False
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("NEW") = "YES" Then
            CMD_NEW.Enabled = True
        Else
            CMD_NEW.Enabled = False
        End If
    End If
    cmd_modify.Enabled = False
    cmd_print.Enabled = True
    cmd_commit.Enabled = False
    If Not USRNAME_RS.EOF Then
        If USRNAME_RS("VIEW") = "YES" Then
            cmd_view.Enabled = True
        Else
            cmd_view.Enabled = False
        End If
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
'    CMD_NEW.Enabled = False
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = True
'    cmd_view.Enabled = True
'
'ElseIf cap = "INI" Then
'    cmd_cancel.Enabled = True
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
'    cmd_cancel.Enabled = True
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
End Sub

Private Sub cmd_commit_Click()
Dim I As Integer
If Trim(txt_modcode.Text) = "PF" Then
    If Len(Trim(TXT_PFTIMPRESSION.Text)) = 0 Then
        MsgBox "Impression Field is Empty, You can not Commit the Report.", vbCritical
        TXT_PFTIMPRESSION.SetFocus
        Exit Sub
    End If
End If
I = MsgBox(" Do you want to Commit the Record?", vbYesNo, vbInformation)
If I = vbYes Then
On Error GoTo MSG:
 
''********SECURITY***************************
'    CURUSRDET_QR(0) = Trim(cmb_finalby.Text)
'    CURUSRDET_QR(0) = Trim(cmb_enterby.Text)
'    CURUSRDET_RS.MoreResults
'    CURUSRDET_RS.Requery
'    If CURUSRDET_RS.RowCount > 0 Then
'        If LOGINTYP > CURUSRDET_RS("usrtype") Then
'          '   no action as he can commit
'        ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
'            MsgBox "Sorry! This Report needs  is to Commited by a Higher Level than  that of Yours"
'            cmd_commit.SetFocus
'            Exit Sub
'        ElseIf LOGINTYP = CURUSRDET_RS("usrtype") Then
'            If Trim(LOGINUSRNAME) <> Trim(CURUSRDET_RS("username")) Then
'                  If LOGINUSRID <> curusrdetrs("usr_id") Then
'                MsgBox "Sorry! This Report is not to be Finalized by You"
'                cmd_commit.SetFocus
'                Exit Sub
'            End If
'        End If
'        If CURUSRDET_RS("COMMIT") <> "YES" Then
'            MsgBox "Sorry! You are not authorised to Commit the Report of this Modality"
'            cmd_commit.SetFocus
'            Exit Sub
'        End If
'    Else
'        MsgBox "Sorry! The Security details of this user is not available check up in DISLNKUSR"
'        cmd_commit.SetFocus
'        Exit Sub
'    End If
''********SECURITY***************************
    
    
 
 
 rdb.BeginTrans
    updrrequisitn1_qr(0) = UCase(cmb_finalby.Text)
'    updrrequisitn1_qr(0) = UCase(cmb_enterby.Text)
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
    ' updating interfaces
    If SELDHEMCYTO_RS.EOF = False Then
    If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("HEM")) = "YES" Then
      updhinterface_qr(0) = Mid(req_chglen(cmb_reqno.Text), 6, 13) 'req_chglen(cmb_reqno.Text)
      updhinterface_qr.Execute
                 
    End If
    If Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
      updcinterface_qr(0) = req_chglen(cmb_reqno.Text)
      updcinterface_qr.Execute
    End If
    End If
    UPDRTIMEC_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
    UPDRTIMEC_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                    
    UPDRTIMEC_QR(2) = req_chglen(cmb_reqno.Text)
    UPDRTIMEC_QR.Execute
    
    '''
    
    MsgBox "Record is Committed Successfully."
End If
rdb.CommitTrans
btnenabl ("INI")

cmd_print.Enabled = True

'********** Shraddha Pandit 22/09/09 **********
  txt_NEW = NEWFLAG
  txt_VIEW = viewflag
  txt_SAVE = flagsave
'**********************************************

Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
End Sub

Private Sub CMD_EMR_Click()
Dim VAR1

'***** Original *****
'VAR1 = MaskEdBox1.FormattedText
'VAR = Shell("C:\EMR1.EXE " & VAR1, vbNormalFocus)
'AppActivate VAR
If Len(Trim(MaskEdBox1.Text)) > 0 Then
    VAR1 = MaskEdBox1.FormattedText
    Dim strLink As String
    'strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
    strLink = "http://10.80.1.9/emrtmhact/TOPBAR.ASP?CASE_NO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If
'*****

'VAR1 = MaskEdBox1.FormattedText
'Dim strLink As String
''strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
'strLink = "http://10.1.4.9/emrweb/Haematopathology.asp?CASENO=" & VAR1 & ""
'Shell "c:\program files\internet explorer\iexplore " & strLink
'End Sub

'***** 20/04/11 *****
'
'Public Sub OpenWebsite(URL As String)
'Dim ret&
'ret& = ShellExecute(Me.hwnd, "Open", URL, vbNullString, vbNullString, SW_SHOW)
End Sub

'********************

Private Sub cmd_mdc_Click()

'Dim b As Integer
'b = InStr(1, cmb_reqno, "$")
'mod_reqno = req_chglen(Mid(Trim(cmb_reqno), 1, b - 1)) 'put change value to this var so that it used later on

'Dim mod_reqno As Variant

'mod_reqno = req_chglen(cmb_reqno.Text)

UPDRREQ_QR(0) = "T"
UPDRREQ_QR(1) = ""
UPDRREQ_QR(2) = req_chglen(cmb_reqno.Text)
UPDRREQ_QR.Execute

cmd_commit.Enabled = True
cmd_mdc.Enabled = False
cmd_modify.Enabled = True

'********* 13/10/08 Shraddha Pandit *********
Dim str As String
Dim UPDTIME_QR As rdoQuery
str = "UPDATE RISLIB.RTIMELOG SET ADRREPDATE=?,ADRREPTIME=? WHERE REQNNO='" & req_chglen(cmb_reqno) & "'"

Set UPDTIME_QR = rdb.CreateQuery("", str)
UPDTIME_QR(0) = Format(txt_date, "YYYY-MM-DD")
UPDTIME_QR(1) = Format(txt_time, "HH:MM:SS")

UPDTIME_QR.Execute

'********************************************

End Sub

Private Sub cmd_modify_Click()
Dim str As String
Dim D1 As String
Dim D2 As String


flagsave = 2
btnenabl ("NEW")
diffcountflag = False
    FLDENABL ("MOD")
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
If txt_1.Visible = True Then
txt_1.Enabled = True
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
ENABL ("ENA")
cmb_enterby.Enabled = True
cmb_finalby.Enabled = True
regcbo.Enabled = True

Command1.Visible = True
'********** Shraddha Pandit 22/09/09 **********
  txt_NEW = NEWFLAG
  txt_VIEW = viewflag
  txt_SAVE = flagsave
  '**********************************************

End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
On Error GoTo MSG:
btnenabl ("NEW")
field_clear
'TXT_CLEAR
txt_1.Text = ""

Command1.Visible = False
cmd_mdc.Enabled = False

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
    If Trim(cmb_enterby.List(0)) = cmb_finalby.List(L1) Then
        cmb_finalby.Text = cmb_finalby.List(L1)
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
If sel_deptrs.RowCount > 0 Then
    sel_deptrs.MoveFirst
    While Not sel_deptrs.EOF

        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        'If (Trim(USRDEPT) = "FED" Or Trim(USRDEPT) = "FEC") And Trim(sel_deptrs("dept_code")) = "FED" Then
    'cmb_dept.Text = Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
     '     End If
        sel_deptrs.MoveNext
        Wend
   End If
Opt_patient.Value = True
    

flagsave = 1
viewflag = 0
NEWFLAG = True
FLDENABL ("ENA")
cmb_dept.SetFocus
MaskEdBox1.Enabled = True
ENABL ("ENA")
TCLEAR
txt_comments.Text = ""

'********** Shraddha Pandit 22/09/09 **********
  txt_NEW = NEWFLAG
  txt_VIEW = viewflag
  txt_SAVE = flagsave
'**********************************************


Exit Sub
MSG:
MsgBox Err.Description
End Sub

'******* Added on 24/09/09 by Shraddha Pandit *******
Public Sub cancelCD34()
btnenabl ("INI")
field_clear
TCLEAR

txt_comments.Text = ""
CMB_MORPHO.Text = ""
txt_comments.Enabled = False
viewflag = False
NEWFLAG = False
Command1.Visible = False
diffcountflag = False

txt_crc.Text = ""
txt_rpi.Text = ""
txt_empc.Text = ""
TXT_LP.Text = ""
ADDENDUMFLAG = False
cmbmpo.Text = ""
End Sub
'*****************************************************

Private Sub cmd_cancel_Click()
btnenabl ("INI")
field_clear
TCLEAR

txt_comments.Text = ""
CMB_MORPHO.Text = ""
txt_comments.Enabled = False
viewflag = False
NEWFLAG = False
Command1.Visible = False
diffcountflag = False

txt_crc.Text = ""
txt_rpi.Text = ""
txt_empc.Text = ""
TXT_LP.Text = "" 'Shraddha Pandit 17/10/08
ADDENDUMFLAG = False 'Shraddha Pandit 30/03/09
cmbmpo.Text = ""
txt_1.Text = ""
'********** Shraddha Pandit 22/09/09 **********
  txt_NEW = NEWFLAG
  txt_VIEW = viewflag
  txt_SAVE = flagsave
  '**********************************************

End Sub

Public Sub field_clear()
Dim obj As Object
cmb_dept.ListIndex = -1
cmb_modality.ListIndex = -1
txt_deptcode.Text = " "
cmb_reqno.ListIndex = -1
cmb_reqno.clear
If Opt_patient.Value = True Then
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
'MaskEdBox1.Mask = ">??/#####"
End If
If opt_staff.Value = True Then
MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
'MaskEdBox1.Mask = "#######"                ' COMMENTED BY VARINDER
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
txt_procedure.Text = " "


'cmb_enterby.Text = "'"

'cmb_finalby.Text = ""
regcbo.Text = ""
TXT_DNACLOIDY.Text = ""
TXT_DNAINDEX.Text = ""
End Sub




Public Sub FLDENABL(EN As String)
If EN = "DIS" Then
        cmb_dept.Enabled = False
        cmb_modality.Enabled = False
        MaskEdBox1.Enabled = False
        cmb_reqno.Enabled = False
        TXT_PATH.Enabled = False
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
        txt_procedure.Enabled = False
        'txt_observed.Visible = False
        txt_comments.Enabled = False
        cmb_enterby.Enabled = False
        cmb_finalby.Enabled = False
        regcbo.Enabled = False
        
        cmb_addendumby.Enabled = False
        txt_addendum.Enabled = False
End If
If EN = "ENA" Then
        cmb_dept.Enabled = True
        cmb_modality.Enabled = True
        MaskEdBox1.Enabled = True
        cmb_reqno.Enabled = True
        TXT_PATH.Enabled = False
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
        txt_procedure.Enabled = False
        'txt_observed.Visible = False
        txt_comments.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
        regcbo.Enabled = True
        
        cmb_addendumby.Enabled = False
        txt_addendum.Enabled = False
End If
If EN = "MOD" Then
        cmb_dept.Enabled = True
        cmb_modality.Enabled = True
        MaskEdBox1.Enabled = True
        cmb_reqno.Enabled = True
        TXT_PATH.Enabled = True
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
        txt_procedure.Enabled = True
        'txt_observed.Visible = True
        txt_comments.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
        regcbo.Enabled = True
        
        'cmb_addendumby.Enabled = True
        'txt_addendum.Enabled = False
        
End If
If EN = "ADDENDUM" Then
    cmb_dept.Enabled = False
    cmb_modality.Enabled = False
    MaskEdBox1.Enabled = False
    cmb_reqno.Enabled = False
    TXT_PATH.Enabled = False
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
    txt_procedure.Enabled = False
    'txt_observed.Visible = False
    'TXT_COMMENTS.Enabled = False
    cmb_enterby.Enabled = False
    cmb_finalby.Enabled = False
    cmb_addendumby.Enabled = False
    txt_addendum.Enabled = False
    If USRNAME_RS("ADDENDUM") = "YES" Then
        txt_addendum.Enabled = True
        cmb_addendumby.Enabled = True
    Else
        txt_addendum.Enabled = False
        cmb_addendumby.Enabled = False
    End If
End If

End Sub
Private Sub cmd_print_Click()
'On Error GoTo MSG:

Dim str1 As String
Dim str2 As String
Dim STR3 As String

Dim REP1 As Boolean
Dim REP2 As Boolean

Dim I As Integer
' On Error GoTo prtmsg:

DoEvents

Dim VAR_REQNNO As String
REP1 = False
REP2 = False
REP3 = False

'***************** 18/11/08 Shraddha Pandit *****************
If SSTab1.Tab = 5 Then
    If Opt_patient = True Then
        Dim sql As String
        Dim SEL_SC_QR As rdoQuery
        Dim SEL_SC_RS As rdoResultset
        Dim rno As String
        Dim UI As String

        UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1

        'sql = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
        'Set SEL_SC_QR = rdb.CreateQuery("", sql)

        rno = req_chglen(cmb_reqno.Text)
        'SEL_SC_QR(0) = rno

       ' Set SEL_SC_RS = SEL_SC_QR.OpenResultset(1, 3)
      '  SEL_SC_RS.MoveFirst
      '  While Not SEL_SC_RS.EOF
       '     If SEL_SC_RS("PRO_CODE") = "U706" Then

                CrystalReport6.SelectionFormula = "{RREQUISITN.REQNNO} = '" & rno & "'"
                CrystalReport6.ReportFileName = "c:\DISREPORTSECLF\DISAUTOCBC_ESR.rpt"

                CrystalReport6.Formulas(10) = "signature='" & Trim(cmb_finalby.Text) & "'"
                
                CrystalReport6.Formulas(0) = "REQNO='" & Trim(cmb_reqno.Text) & "'"
                CrystalReport6.Formulas(1) = "PROVDIA = '" & txt_procedure & "'"
                CrystalReport6.Formulas(2) = "NAM = '" & txt_name & "'"
                CrystalReport6.Formulas(3) = "Age = '" & txt_age & "'"
                CrystalReport6.Formulas(4) = "Sex = '" & txt_sex & "'"
                CrystalReport6.Formulas(5) = "USRID = '" & UI & "'"
                CrystalReport6.Formulas(17) = "signature='" & Trim(cmb_finalby.Text) & "'"
                CrystalReport6.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport6.WindowState = crptMaximized
                CrystalReport6.Action = 1
         '   End If
        '    SEL_SC_RS.MoveNext
      '  Wend
        Exit Sub
    End If
    If opt_staff = True Then
        Dim ssql As String
        Dim sSEL_SC_QR As rdoQuery
        Dim sSEL_SC_RS As rdoResultset
        Dim srno As String
        Dim sUI As String

        sUI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1

     '   ssql = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
     '   Set sSEL_SC_QR = rdb.CreateQuery("", ssql)

        srno = req_chglen(cmb_reqno.Text)
      '  sSEL_SC_QR(0) = srno

     '   Set sSEL_SC_RS = sSEL_SC_QR.OpenResultset(1, 3)
      '  sSEL_SC_RS.MoveFirst
      '  While Not sSEL_SC_RS.EOF
      '      If sSEL_SC_RS("PRO_CODE") = "U706" Then

                CrystalReport6.SelectionFormula = "{RREQUISITN.REQNNO} = '" & srno & "'"
                CrystalReport6.ReportFileName = "c:\DISREPORTSECLF\DISAUTOCBC_ESR_staff.rpt"

                CrystalReport6.Formulas(0) = "REQNO='" & Trim(cmb_reqno.Text) & "'"
                CrystalReport6.Formulas(1) = "PROVDIA = '" & txt_procedure & "'"
                CrystalReport6.Formulas(2) = "USRID = '" & sUI & "'"
                CrystalReport6.Formulas(3) = "Nam = '" & txt_name & "'"
                CrystalReport6.Formulas(4) = "Age = '" & txt_age & "'"
                CrystalReport6.Formulas(5) = "Sex = '" & txt_sex & "'"
                
                CrystalReport6.Formulas(10) = "ST = 'ST'"
                
                CrystalReport6.Formulas(17) = "signature='" & Trim(cmb_finalby.Text) & "'"
                CrystalReport6.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport6.WindowState = crptMaximized
                CrystalReport6.Action = 1
        '    End If
      '      sSEL_SC_RS.MoveNext
      '  Wend
        Exit Sub
    End If
End If
'************************************************************


'-------------- Shraddha Pandit 20/11/08 --------------------
If Trim(txt_esr.Text) <> "" Then
    If Opt_patient = True Then
        Dim sql1 As String
        Dim SEL_SC_QR1 As rdoQuery
        Dim SEL_SC_RS1 As rdoResultset
        Dim rno1 As String
        Dim UI1 As String

        UI1 = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1

     '   sql1 = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
     '   Set SEL_SC_QR1 = rdb.CreateQuery("", sql1)

        rno1 = req_chglen(cmb_reqno.Text)
     '   SEL_SC_QR1(0) = rno1

      '  Set SEL_SC_RS1 = SEL_SC_QR1.OpenResultset(1, 3)
      '  SEL_SC_RS1.MoveFirst
     '   While Not SEL_SC_RS1.EOF
     '       If SEL_SC_RS1("PRO_CODE") = "U706" Then

                CrystalReport6.SelectionFormula = "{RREQUISITN.REQNNO} = '" & rno1 & "'"
                CrystalReport6.ReportFileName = "c:\DISREPORTSECLF\DISAUTOCBC_ESR.rpt"
                
                CrystalReport6.Formulas(0) = "REQNO='" & Trim(cmb_reqno.Text) & "'"
                CrystalReport6.Formulas(1) = "PROVDIA = '" & txt_procedure & "'"
                CrystalReport6.Formulas(2) = "USRID = '" & UI1 & "'"
                CrystalReport6.Formulas(3) = "Nam = '" & txt_name & "'"
                CrystalReport6.Formulas(4) = "Age = '" & txt_age & "'"
                CrystalReport6.Formulas(5) = "Sex = '" & txt_sex & "'"
                CrystalReport6.Formulas(17) = "signature='" & Trim(cmb_finalby.Text) & "'"
                CrystalReport6.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport6.WindowState = crptMaximized
                CrystalReport6.Action = 1
     '       End If
     '       SEL_SC_RS1.MoveNext
       ' Wend
    End If
    If opt_staff = True Then
        Dim ssql1 As String
        Dim sSEL_SC_QR1 As rdoQuery
        Dim sSEL_SC_RS1 As rdoResultset
        Dim srno1 As String
        Dim sUI1 As String

        sUI1 = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1

      '  ssql1 = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO = ? "
      '  Set sSEL_SC_QR1 = rdb.CreateQuery("", ssql1)

        srno1 = req_chglen(cmb_reqno.Text)
      '  sSEL_SC_QR1(0) = srno1

      '  Set sSEL_SC_RS1 = sSEL_SC_QR1.OpenResultset(1, 3)
      '  sSEL_SC_RS1.MoveFirst
      '  While Not sSEL_SC_RS1.EOF
      '      If sSEL_SC_RS1("PRO_CODE") = "U706" Then

                CrystalReport6.SelectionFormula = "{RREQUISITN.REQNNO} = '" & srno1 & "'"
                CrystalReport6.ReportFileName = "c:\DISREPORTSECLF\DISAUTOCBC_ESR_staff.rpt"
                
                CrystalReport6.Formulas(0) = "REQNO='" & Trim(cmb_reqno.Text) & "'"
                CrystalReport6.Formulas(1) = "PROVDIA = '" & txt_procedure & "'"
                CrystalReport6.Formulas(2) = "USRID = '" & sUI1 & "'"
                CrystalReport6.Formulas(3) = "Nam = '" & txt_name & "'"
                CrystalReport6.Formulas(4) = "Age = '" & txt_age & "'"
                CrystalReport6.Formulas(5) = "Sex = '" & txt_sex & "'"
                CrystalReport6.Formulas(17) = "signature='" & Trim(cmb_finalby.Text) & "'"
                CrystalReport6.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport6.WindowState = crptMaximized
                CrystalReport6.Action = 1
     '       End If
     '       sSEL_SC_RS1.MoveNext
      '  Wend
    End If
End If
'------------------------------------------------------------
If SELDHEMCYTO_RS.EOF = False Then
    If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("HEM")) = "YES" Or Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
        REP1 = True
    End If
    'If Trim(SELDHEMCYTO_RS("CYTO")) = "YES" Or Trim(SELDHEMCYTO_RS("SMAR")) = "YES" Then
        REP2 = True
    'End If
    If Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
        REP3 = True
    End If
    
    '***** Shraddha 06/05/2011 *****
    If Trim(SELDHEMCYTO_RS("MPO")) = "YES" Then
        REP1 = True
    End If
    End If
    '*******************************
          
    VAR_REQNNO = req_chglen(cmb_reqno.Text)
    If REFDRFLAG1 = False Then
        REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
        REFDR_QR(1) = Trim(req_chglen(cmb_reqno.Text))
        Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        REFDRFLAG1 = True
    Else
        REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
        REFDR_QR(1) = Trim(req_chglen(cmb_reqno.Text))
        REFDR_RS.MoreResults
        REFDR_RS.Requery
    End If

    If SELRREQUISITNFLAG1 = False Then
        SELRREQUISITN_QR(0) = Trim(MaskEdBox1.FormattedText)
        SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
        SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
        SELRREQUISITN_QR(3) = "T"
        SELRREQUISITN_QR(4) = "F"
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG1 = True
    Else
        SELRREQUISITN_QR(0) = Trim(MaskEdBox1.FormattedText)
        SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
        SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
        SELRREQUISITN_QR(3) = "T"
        SELRREQUISITN_QR(4) = "F"
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    '            If SELRREQUISITN_Rs.RowCount > 0 Then
    '                prepath = ""
    '                While Not SELRREQUISITN_Rs.EOF
    '                    If Trim(SELRREQUISITN_Rs("reqnno")) <> req_chglen(Trim(cmb_reqno.Text)) Then
    '                        prepath = prepath & req_chglen(SELRREQUISITN_Rs("REQNNO")) & " "
    '                    End If
    '
    '                    SELRREQUISITN_Rs.MoveNext
    '                Wend
    '             End If
    '
    If Opt_patient.Value = True Then
        '''''''''''''''kanchan
        'If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Then
        
        'If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("MPO")) = "YES" Then
        If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("MPO")) = "YES" Then
            '***** Shraddha 09/03/11 *****
            '***** Shraddha 10/12/10 *****
        'str = "SELECT * FROM DISLIB.HINTERFACE WHERE REQNO=?"
        SEL_DHINTER_QR(0) = VAR_REQNNO
        Set SEL_DHINTER_RS = SEL_DHINTER_QR.OpenResultset(1, 3)
        If Not SEL_DHINTER_RS.EOF Then
        '*****************************
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
        For SH = 0 To 35
         CrystalReport1.Formulas(SH) = ""
        Next
        'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' "
        CrystalReport1.SelectionFormula = "{reporthlf1.REQNO} = '" & VAR_REQNNO & "' "
        'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISAUTOCBC.RPT"
        'Commented and added by Shraddha on 13/09/10
        
        CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISAUTOCBC.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "F" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL AUTOMATED COMPLETE BLOOD COUNT REPORT" & "'"
            str1 = ""
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "AUTOMATED COMPLETE BLOOD COUNT REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        
        '***** Shraddha 04/01/11 *****
                    
            sel_dhemcbc_qr(0) = Trim(VAR_REQNNO)
            sel_dhemcbc_qr(1) = "R"
            Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 1)
            
            Dim strRBC, strWBC, strPLATELET As String
            strRBC = ""
            strWBC = ""
            strPLATELET = ""
            While Not sel_dhemcbc_rs.EOF
                'strRBC = strRBC & Trim(SEL_DHEMCBC_RS("PRO_CODE")) & ":" & Trim(SEL_DHEMCBC_RS("OBSERVALUE")) & " "
                strRBC = strRBC & Trim(sel_dhemcbc_rs("TESTNAME")) & ":" & Trim(sel_dhemcbc_rs("OBSERVALUE")) & " "
                sel_dhemcbc_rs.MoveNext
            Wend
            
            CrystalReport1.Formulas(13) = "RBC = '" & strRBC & "'"
            
            '*****
            sel_dhemcbc_qr(0) = Trim(VAR_REQNNO)
            sel_dhemcbc_qr(1) = "W"
            Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 1)
                        
            While Not sel_dhemcbc_rs.EOF
                'strWBC = strWBC & Trim(SEL_DHEMCBC_RS("PRO_CODE")) & ":" & Trim(SEL_DHEMCBC_RS("OBSERVALUE")) & " "
                strWBC = strWBC & Trim(sel_dhemcbc_rs("TESTNAME")) & ":" & Trim(sel_dhemcbc_rs("OBSERVALUE")) & " "
                sel_dhemcbc_rs.MoveNext
            Wend
            
            CrystalReport1.Formulas(14) = "WBC = '" & strWBC & "'"
            
            '*****
            sel_dhemcbc_qr(0) = Trim(VAR_REQNNO)
            sel_dhemcbc_qr(1) = "P"
            Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 1)
                        
            While Not sel_dhemcbc_rs.EOF
                'strPLATELET = strPLATELET & Trim(SEL_DHEMCBC_RS("PRO_CODE")) & ":" & Trim(SEL_DHEMCBC_RS("OBSERVALUE")) & " "
                strPLATELET = strPLATELET & Trim(sel_dhemcbc_rs("TESTNAME")) & ":" & Trim(sel_dhemcbc_rs("OBSERVALUE")) & " "
                sel_dhemcbc_rs.MoveNext
            Wend
            
            CrystalReport1.Formulas(15) = "PLATELET = '" & strPLATELET & "'"
        
        '*****************************
        CrystalReport1.Formulas(11) = ""
        CrystalReport1.Formulas(16) = ""
        CrystalReport1.Formulas(18) = ""
        CrystalReport1.Formulas(19) = ""
        CrystalReport1.Formulas(20) = ""
        CrystalReport1.Formulas(21) = ""
        CrystalReport1.Formulas(22) = ""
         CrystalReport1.Formulas(23) = ""
         CrystalReport1.Formulas(24) = ""
         CrystalReport1.Formulas(25) = ""
         
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(17) = "signature='" & Trim(cmb_finalby.Text) & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'        For SH = 0 To 25
'        MsgBox SH & ". " & CrystalReport1.Formulas(26)
'        Next
        CrystalReport1.Action = 1
        
        '***** Haemato *****
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        CrystalReport4.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
        CrystalReport5.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
            
            If Trim(txt_modcode.Text) = "HL" Then
                CrystalReport5.ReportFileName = "C:\DISREPORTSECLF\HEMATO.RPT"
                CrystalReport5.Formulas(0) = "REQNO='" & Trim(cmb_reqno.Text) & "'"
                CrystalReport5.Formulas(17) = "signature='" & Trim(cmb_finalby.Text) & "'"
                CrystalReport5.WindowState = crptMaximized
                CrystalReport5.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport5.Action = 1
            End If
        '***** Haemato *****
        
        
        Else
            '*****************************
        
        
        
            str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
            CrystalReport4.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
            CrystalReport5.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
            
            If Trim(txt_modcode.Text) = "HL" Then
                CrystalReport5.ReportFileName = "C:\DISREPORTSECLF\HEMATO.RPT"
                                
               '*****************************
                
                CrystalReport4.ReportFileName = "C:\DISREPORTSEC\DISAUTOCBC.rpt"
                CrystalReport4.Formulas(0) = "REQNO='" & Trim(cmb_reqno.Text) & "'"
                CrystalReport4.Formulas(20) = "PROFIN = 'AUTOMATED COMPLETE BLOOD COUNT REPORT'"
                                
                CrystalReport5.Formulas(0) = "REQNO='" & Trim(cmb_reqno.Text) & "'"
                          
           
                '***** 29/04/11 *****
                CrystalReport4.Formulas(17) = "signature='" & Trim(cmb_finalby.Text) & "'"
                CrystalReport5.Formulas(17) = "signature='" & Trim(cmb_finalby.Text) & "'"
                '********************
           
               CrystalReport4.WindowState = crptMaximized
               CrystalReport5.WindowState = crptMaximized
               CrystalReport4.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
               CrystalReport5.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
               CrystalReport4.Action = 1
               CrystalReport5.Action = 1
            End If 'HL
        End If  'CBC/HEM/CG
        End If '09/03/11
        'If Trim(SELDHEMCYTO_RS("HEM")) = "YES" Then
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        CrystalReport5.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
     
        If Trim(txt_modcode.Text) = "HL" Then
            CrystalReport5.ReportFileName = "C:\DISREPORTSECLF\HEMATO.RPT"
            CrystalReport4.ReportFileName = "C:\DISREPORTSEC\DISAUTOCBC.rpt"
            CrystalReport5.Formulas(0) = "REQNO='" & Trim(cmb_reqno.Text) & "'"
            CrystalReport5.Formulas(1) = "PROVDIA = '" & txt_procedure & "'"
            CrystalReport5.Formulas(17) = "signature='" & Trim(cmb_finalby.Text) & "'"
            CrystalReport5.WindowState = crptMaximized
            'CrystalReport5.Destination = crptToPrinter
            CrystalReport5.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
           'CrystalReport5.Action = 1
        End If 'HL
     
        ' End If
 
        '''  NEW
 
        If Trim(SELDHEMCYTO_RS("CYTO")) = "" Or Trim(SELDHEMCYTO_RS("SMAR")) = "YES" Then
        For SH = 0 To 35
         CrystalReport1.Formulas(SH) = ""
        Next
            
            str = ipadd '& " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
            CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
                
            If Trim(txt_modcode.Text) = "HL" Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShlN1A.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
                CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
                    
                If Trim(REFDR_RS("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "PERIPHERAL BLOOD REPORT" & "'"
                    str1 = ""
                    str1 = "Entered By " & Trim(cmb_enterby.Text)
                    If Len(Trim(cmb_enterby.Text)) > 0 Then
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                    Else
                        CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
                    End If
                Else
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "PERIPHERAL BLOOD REPORT" & "'"
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                End If
                CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(26) = "RGT = '" & time1 & "'"
                If Trim(txt_lapscore.Text) = "" And Trim(TXT_MPO.Text) = "" And Trim(cmb_nse.Text) = "" And Trim(TXT_PAS.Text) = "" And Trim(TXT_NSE.Text) = "" And Trim(TXT_TRAP.Text) = "" Then
                        'If CYTOCHEMFLAG = True Then
                        '                 CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
                        '           Else
                        '            CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
                        'End If
                End If
                If Trim(txt_inr.Text) = "" And Trim(TXT_PT(0).Text) = "" And Trim(txt_ptctl.Text) = "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_TT.Text) = "" And Trim(TXT_FIBRINOGEN.Text) = "" And Trim(TXT_FDPDMIER.Text) = "" And Trim(TXT_OTHERS.Text) = "" Then
                    CrystalReport1.Formulas(11) = "CP = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(11) = "CP = '" & "ABC" & "'"
                End If

                If Trim(TXT_NORMOBLASTS.Text) = "" And Trim(TXT_OTHERCELLSRBC.Text) = "" And Trim(txt_rbcmorphology.Text) = "" And Trim(txt_reticcount.Text) = "" And Trim(txt_malarialparasite.Text) = "" And Trim(txt_esr.Text) = "" Then  '18/11/08
                    CrystalReport1.Formulas(15) = "RBC = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(15) = "RBC = '" & "ABC" & "'"
                End If
                    
                'If Trim(TXT_MORPHOLOGY.Text) = "" And Trim(TXT_ADEQUACY.Text) = "" Then
                If Trim(CMB_MORPHO.Text) = "" And Trim(TXT_ADEQUACY.Text) = "" Then
                    CrystalReport1.Formulas(16) = "PLATELETS = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(16) = "PLATELETS = '" & "ABC" & "'"
                End If
                    
                If Trim(txt_prolymphocytes.Text) = "" And Trim(txt_Sezarycells.Text) = "" And Trim(txt_hairycells.Text) = "" And Trim(txt_othercells.Text) = "" Then
                    CrystalReport1.Formulas(17) = "prolyn = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(17) = "prolyn = '" & "123" & "'"
                End If
                    
                If Trim(txt_blasts.Text) = "" And Trim(TXT_PROMYELOCYTESWBC.Text) = "" And Trim(txt_myelocytes.Text) = "" And Trim(txt_metamyelocytes.Text) = "" And Trim(txt_plasmacells.Text) = "" And Trim(txt_alymphocytes.Text) = "" Then
                    CrystalReport1.Formulas(18) = "blast = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(18) = "blast = '" & "123" & "'"
                End If
                    
                If Trim(txt_mneutrophils.Text) = "" And Trim(txt_mlymphocytes.Text) = "" And Trim(txt_mmonocytes.Text) = "" And Trim(txt_meosinophils.Text) = "" And Trim(txt_mbasophills.Text) = "" Then
                    CrystalReport1.Formulas(19) = "cbcm = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(19) = "cbcm = '" & "123" & "'"
                End If
                    
                If Trim(TXT_HCT.Text) = "" And Trim(TXT_MCH.Text) = "" And Trim(TXT_MCV.Text) = "" And Trim(TXT_MCHC.Text) = "" And Trim(TXT_MPV.Text) = "" And Trim(TXT_RDW.Text) = "" And Trim(TXT_ABNORMAL.Text) = "" Then
                    CrystalReport1.Formulas(20) = "rbci = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(20) = "rbci = '" & "123" & "'"
                End If
                    
                If Trim(txt_neutrophils.Text) = "" And Trim(txt_lymphocytes.Text) = "" And Trim(txt_monocytes.Text) = "" And Trim(txt_eosinophils.Text) = "" And Trim(txt_basophils.Text) = "" And Trim(txt_immaturegranulocytes.Text) = "" Then
                    CrystalReport1.Formulas(21) = "dlc = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(21) = "dlc = '" & "123" & "'"
                End If
                    
                If Trim(TXT_HB.Text) = "" And Trim(TXT_RBC.Text) = "" And Trim(TXT_TWBC.Text) = "" And Trim(txt_platelets.Text) = "" Then
                    CrystalReport1.Formulas(22) = "cbcc = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(22) = "cbcc = '" & "123" & "'"
                End If
                    
                If Trim(SELDHEMCYTO_RS("smar")) = "YES" Then
                    CrystalReport1.Formulas(23) = "mar = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(23) = "mar = '" & "123" & "'"
                End If
                    
                If Trim(txt_comments.Text) <> "" Then
                    CrystalReport1.Formulas(24) = "nocomments='" & "" & "'"
                    CrystalReport1.Formulas(25) = "comments='" & "COMMENTS:" & "'"
                Else
                    If Trim(txt_1.Text) <> "" Then
                     CrystalReport1.Formulas(24) = "nocomments='" & "" & "'"
                    CrystalReport1.Formulas(25) = "comments='" & "COMMENTS:" & "'"
                    Else
                     CrystalReport1.Formulas(24) = "nocomments='" & "YES" & "'"
                    CrystalReport1.Formulas(25) = "comments='" & "" & "'"
                    End If
                
                   
                End If
                
                If Trim(txt_addendum.Text) <> "" Then
                    CrystalReport1.Formulas(26) = "addedndum='ADDENDUM:'"
                End If
                ' PRINT REPPORT
           
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
        End If  'CYTO/SURFACE MARKER
        ''** FOR COAGULATION PROFILE
        If Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
            CrystalReport3.Formulas(0) = "REQNO='" & cmb_reqno.Text & "'"
            str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
            CrystalReport3.Formulas(1) = "USRID='" & str & "'"
            CrystalReport3.Formulas(2) = "PROFIN='" & "HAEMATO PATHOLOGY  REPORT" & "'"
    
            If Trim(txt_comments.Text) <> "" Then
                CrystalReport3.Formulas(3) = "nocomments='" & "" & "'"
                CrystalReport3.Formulas(4) = "COMMENTS='" & "COMMENTS:" & "'"
            Else
                CrystalReport3.Formulas(3) = "nocomments='" & "YES" & "'"
                CrystalReport3.Formulas(4) = "COMMENTS='" & "" & "'"
            End If
            
            If Trim(txt_impression.Text) <> "" Then
                CrystalReport3.Formulas(5) = "noimpress='" & "" & "'"
                CrystalReport3.Formulas(6) = "impress='" & "IMPRESSION:" & "'"
            Else
                CrystalReport3.Formulas(5) = "noimpress='" & "YES" & "'"
                CrystalReport3.Formulas(6) = "impress='" & "" & "'"
            End If
            
            If Trim(TXT_FDPDMIER.Text) <> "" Then
                CrystalReport3.Formulas(7) = "D-DIMER='" & "D-Dimer test is performed manually by Semi-Quantitative Latex Agglutination Technique." & "'"
            End If
            
            If Trim(TXT_TT.Text) <> "" Then
                CrystalReport3.Formulas(8) = "TT='" & "TT test is performed manually." & "'"
            End If
            
            'Dim STR As String
            ' STR = ""
            '            If Trim(TXT_PT(0).Text) <> "" Then
            '                STR = "PT"
            '
            '            End If
            '            If Trim(TXT_PT(0).Text) <> "" And Trim(TXT_PTTK.Text) <> "" Then
            '                STR = STR & "/APTT"
            '            ElseIf Trim(TXT_PT(0).Text) = "" And Trim(TXT_PTTK.Text) <> "" Then
            '                STR = STR & "APTT"
            '            End If
            
            If Trim(TXT_PT(0).Text) <> "" And Trim(TXT_PTTK.Text) <> "" And Trim(TXT_FIBRINOGEN.Text) <> "" Then
                str = "PT/APTT/FIBRINOGEN"
            ElseIf Trim(TXT_PT(0).Text) = "" And Trim(TXT_PTTK.Text) <> "" And Trim(TXT_FIBRINOGEN.Text) <> "" Then
                str = "APTT/FIBRINOGEN"
            ElseIf Trim(TXT_PT(0).Text) = "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_FIBRINOGEN.Text) <> "" Then
                str = "FIBRINOGEN"
            ElseIf Trim(TXT_PT(0).Text) <> "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_FIBRINOGEN.Text) <> "" Then
                str = "PT/FIBRINOGEN"
            ElseIf Trim(TXT_PT(0).Text) <> "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_FIBRINOGEN.Text) = "" Then
                str = "PT"
            ElseIf Trim(TXT_PT(0).Text) = "" And Trim(TXT_PTTK.Text) <> "" And Trim(TXT_FIBRINOGEN.Text) = "" Then
                str = "APTT"
            ElseIf Trim(TXT_PT(0).Text) <> "" And Trim(TXT_PTTK.Text) <> "" And Trim(TXT_FIBRINOGEN.Text) = "" Then
                str = "PT/APTT"
            End If
            
'            If Option1.Value = True Then
'                str1 = "ACL 100"
'            ElseIf Option2.Value = True Then
'                str1 = "ACL 200"
'            ElseIf Option3.Value = True Then
'             str1 = "ELP 01"
'            Else
            
             str1 = cmb_machname.Text
'            End If
            
            If Trim(TXT_PT(0).Text) <> "" Or Trim(TXT_PTTK.Text) <> "" Or Trim(TXT_FIBRINOGEN.Text <> "") Then
                CrystalReport3.Formulas(9) = "PT='" & str & " tests are performed on automated coagulometer " & str1 & ".'"
            End If
            '            If Trim(TXT_PT(0).Text) <> "" And Trim(TXT_PTTK.Text) <> "" Then
            '                CrystalReport3.Formulas(9) = "PT='" & "PT/APTT tests are performed on ACL automated coagulometer." & "'"
            '            End If
            '            If Trim(TXT_PT(0).Text) <> "" And Trim(TXT_FIBRINOGEN.Text <> "") Then
            '                CrystalReport3.Formulas(9) = "PT='" & "PT/APTT/Fibrinogen tests are performed on ACL automated coagulometer." & "'"
            '            End If
            
            CrystalReport3.SelectionFormula = "{RREQUISITN.REQNNO}='" & VAR_REQNNO & "'"
            'CrystalReport3.SelectionFormula = "{REPORTCPLF.REQNO}='" & VAR_REQNNO & "'"
            'CrystalReport3.ReportFileName = "C:\DISREPORTSEC\DISCPDET.rpt"
            'Commented and added by Shraddha on 14/09/10
            'CrystalReport3.ReportFileName = "C:\DISREPORTSECLF\DISCPDETlf.rpt"
            CrystalReport3.ReportFileName = "C:\DISREPORTSECLF\DISCPDET_INTER.rpt"
            CrystalReport3.Formulas(17) = "signature='" & Trim(cmb_finalby.Text) & "'"
            'CrystalReport3.Destination = crptToPrinter
            CrystalReport3.Destination = crptToWindow
            CrystalReport3.WindowState = crptMaximized
            CrystalReport3.Connect = "ODBC;Database=DISLIB,PATOOLSLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
            CrystalReport3.Action = 1
        End If
  
        ''   NEW
 
        If REP1 = True And REP2 = True Then
                 'MsgBox "There are two reports print each one of them"
        End If
        If REP3 = True And REP2 = True Then
            'MsgBox "There are two reports print each one of them"
        End If
        If REP1 = True And REP2 = True And REP3 = True Then
            'MsgBox "There are THREE reports print each one of them"
        End If
    End If

    '' HL ADDENDUM
    If ADDENDUMFALG = True Then
        CrystalReport2.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        CrystalReport1.Formulas(1) = "PROFIN = '" & "PERIPHERAL BLOOD REPORT" & "'"
        CrystalReport2.ReportFileName = "c:\DISREPORTSEC\DISHLADD.RPT"
        CrystalReport2.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
        CrystalReport2.WindowState = crptMaximized
        CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport2.Action = 1
        MsgBox "Print also the Addendum Report"
    End If
    
    If opt_staff.Value = True Then
        str = ipadd '& " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
        If Trim(txt_modcode.Text) = "EK" Then
            CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEHLASTAFF.RPT"
            'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISAUTOCBCS.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
            '        If Trim(prepath) = "" Then
            '            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            '        Else
            '            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
            '        End If
            CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
            
            If Trim(REFDR_RS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & Trim(cmb_enterby.Text)
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        End If
        
        If Trim(txt_modcode.Text) = "HL" Then
            '' NEW CODES FOR STAFF REPORTS
            'If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("HEM")) = "YES" Or Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
            
            '***** Shraddha 09/03/11 *****
        'str = "SELECT * FROM DISLIB.HINTERFACE WHERE REQNO=?"
        SEL_DHINTER_QR(0) = VAR_REQNNO
        Set SEL_DHINTER_RS = SEL_DHINTER_QR.OpenResultset(1, 3)
        If Not SEL_DHINTER_RS.EOF Then
        '*****************************
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
        
        'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' "
        CrystalReport1.SelectionFormula = "{reporthlf1.REQNO} = '" & VAR_REQNNO & "' "
        'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISAUTOCBC.RPT"
        'Commented and added by Shraddha on 13/09/10
        CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISAUTOCBC.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "F" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
            str1 = ""
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        
        '***** Shraddha 04/01/11 *****
                    
            sel_dhemcbc_qr(0) = Trim(VAR_REQNNO)
            sel_dhemcbc_qr(1) = "R"
            Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 1)
            
            'Dim strRBC, strWBC, strPLATELET As String
            strRBC = ""
            strWBC = ""
            strPLATELET = ""
            While Not sel_dhemcbc_rs.EOF
                'strRBC = strRBC & Trim(SEL_DHEMCBC_RS("PRO_CODE")) & ":" & Trim(SEL_DHEMCBC_RS("OBSERVALUE")) & " "
                strRBC = strRBC & Trim(sel_dhemcbc_rs("TESTNAME")) & ":" & Trim(sel_dhemcbc_rs("OBSERVALUE")) & " "
                sel_dhemcbc_rs.MoveNext
            Wend
            
            CrystalReport1.Formulas(13) = "RBC = '" & strRBC & "'"
            
            '*****
            sel_dhemcbc_qr(0) = Trim(VAR_REQNNO)
            sel_dhemcbc_qr(1) = "W"
            Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 1)
                        
            While Not sel_dhemcbc_rs.EOF
                'strWBC = strWBC & Trim(SEL_DHEMCBC_RS("PRO_CODE")) & ":" & Trim(SEL_DHEMCBC_RS("OBSERVALUE")) & " "
                strWBC = strWBC & Trim(sel_dhemcbc_rs("TESTNAME")) & ":" & Trim(sel_dhemcbc_rs("OBSERVALUE")) & " "
                sel_dhemcbc_rs.MoveNext
            Wend
            
            CrystalReport1.Formulas(14) = "WBC = '" & strWBC & "'"
            
            '*****
            sel_dhemcbc_qr(0) = Trim(VAR_REQNNO)
            sel_dhemcbc_qr(1) = "P"
            Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 1)
                        
            While Not sel_dhemcbc_rs.EOF
                'strPLATELET = strPLATELET & Trim(SEL_DHEMCBC_RS("PRO_CODE")) & ":" & Trim(SEL_DHEMCBC_RS("OBSERVALUE")) & " "
                strPLATELET = strPLATELET & Trim(sel_dhemcbc_rs("TESTNAME")) & ":" & Trim(sel_dhemcbc_rs("OBSERVALUE")) & " "
                sel_dhemcbc_rs.MoveNext
            Wend
            
            CrystalReport1.Formulas(15) = "PLATELET = '" & strPLATELET & "'"
            
            '***** Staff *****
            CrystalReport1.SelectionFormula = "{REPORTHLF1.REQNO} = '" & Trim(VAR_REQNNO) & "'"
            CrystalReport1.ReportFileName = "C:\DisreportsecLF\DISAUTOCBC.RPT"
            
            SELSTAFFINFO_QR(0) = Trim(VAR_REQNNO)
            Set SELSTAFFINFO_RS = SELSTAFFINFO_QR.OpenResultset(2, 3)
            CrystalReport1.Formulas(2) = "NAM = '" & Trim(SELSTAFFINFO_RS("pname")) & "'"
            
            CrystalReport1.Formulas(3) = "staffagesex = '" & Trim(SELSTAFFINFO_RS("AGE")) & "/" & Trim(SELSTAFFINFO_RS("SEX")) & "'"
            
            '*****************
            
        
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
        '***************************** 09/03/11
        Else
            
            
            If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Then
                str = ipadd '& " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
                'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShlN1AASTAFF.RPT"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISAUTOCBCS.RPT"
                'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShlN1AA.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
                CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
                
                If Trim(REFDR_RS("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "AUTOMATED COMPLETE BLOOD COUNT REPORT" & "'"
                    str1 = ""
                    str1 = "Operated by " & Trim(cmb_enterby.Text)
                    
                    If Len(Trim(cmb_enterby.Text)) > 0 Then
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                    Else
                        CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
                    End If
                Else
                    str1 = ""
                    str1 = "Operated by " & Trim(cmb_enterby.Text) 'Trim(cmb_enterby.Text)
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "AUTOMATED COMPLETE BLOOD COUNT REPORT" & "'"
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                End If
                CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
                If Trim(txt_comments.Text) <> "" Then
                    CrystalReport1.Formulas(24) = "nocomments='" & "" & "'"
                    CrystalReport1.Formulas(25) = "comments='" & "COMMENTS:" & "'"
                Else
                    CrystalReport1.Formulas(24) = "nocomments='" & "YES" & "'"
                    CrystalReport1.Formulas(25) = "comments='" & "" & "'"
                End If
                
                '' PRINT REPORT
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
                '           If REP1 = True And REP2 = True Then
                '              MsgBox "There are two reports print this one first"
                '           End If
                '  End If 'HL
                'End If  'CBC/HEM/CG
                '        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
                ''        If Trim(prepath) = "" Then
                ''            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                ''        Else
                ''            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
                ''        End If
                '        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
                '        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
                '            CrystalReport1.Formulas(5) = "PROFIN = '" & "PERIPHERAL BLOOD REPORT" & "'"
                '            STR1 = ""
                '            STR1 = "Entered By " & Trim(cmb_enterby.Text)
                '            If Len(Trim(cmb_enterby.Text)) > 0 Then
                '                CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
                '            Else
                '                 CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
                '            End If
                '         Else
                '
                '            CrystalReport1.Formulas(5) = "PROFIN = '" & "PERIPHERAL BLOOD REPORT" & "'"
                '            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
                '        End If
                '        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
                '
                '          If Trim(txt_lapscore.Text) = "" And Trim(TXT_MPO.Text) = "" And Trim(cmb_nse.Text) = "" And Trim(TXT_PAS.Text) = "" And Trim(TXT_NSE.Text) = "" And Trim(TXT_TRAP.Text) = "" Then
                '           'If CYTOCHEMFLAG = True Then
                '             CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
                '           Else
                '            CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
                '           'End If
                '           End If
                '            If Trim(txt_inr.Text) = "" And Trim(TXT_PT(0).Text) = "" And Trim(txt_ptctl.Text) = "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_TT.Text) = "" And Trim(TXT_FIBRINOGEN.Text) = "" And Trim(TXT_FDPDMIER.Text) = "" And Trim(txt_others.Text) = "" Then
                '
                '             CrystalReport1.Formulas(11) = "CP = '" & "" & "'"
                '           Else
                '            CrystalReport1.Formulas(11) = "CP = '" & "ABC" & "'"
                '
                '           End If
                '
                '           If Trim(TXT_NORMOBLASTS.Text) = "" And Trim(TXT_OTHERCELLSRBC.Text) = "" And Trim(txt_rbcmorphology.Text) = "" And Trim(txt_reticcount.Text) = "" And Trim(txt_malarialparasite.Text) = "" And Trim(TXT_ESRRBC.Text) = "" Then
                '
                '             CrystalReport1.Formulas(15) = "RBC = '" & "" & "'"
                '           Else
                '            CrystalReport1.Formulas(15) = "RBC = '" & "ABC" & "'"
                '
                '           End If
                '           If Trim(txt_morphology.Text) = "" And Trim(TXT_ADEQUACY.Text) = "" Then
                '
                '             CrystalReport1.Formulas(16) = "PLATELETS = '" & "" & "'"
                '           Else
                '            CrystalReport1.Formulas(16) = "PLATELETS = '" & "ABC" & "'"
                '
                '           End If
                '           If Trim(txt_prolymphocytes.Text) = "" And Trim(txt_Sezarycells.Text) = "" And Trim(txt_hairycells.Text) = "" And Trim(txt_othercells.Text) = "" Then
                '                CrystalReport1.Formulas(17) = "prolyn = '" & "" & "'"
                '           Else
                '                CrystalReport1.Formulas(17) = "prolyn = '" & "123" & "'"
                '           End If
                '           If Trim(txt_blasts.Text) = "" And Trim(TXT_PROMYELOCYTESWBC.Text) = "" And Trim(txt_myelocytes.Text) = "" And Trim(txt_metamyelocytes.Text) = "" And Trim(txt_plasmacells.Text) = "" And Trim(txt_alymphocytes.Text) = "" Then
                '                CrystalReport1.Formulas(18) = "blast = '" & "" & "'"
                '           Else
                '                CrystalReport1.Formulas(18) = "blast = '" & "123" & "'"
                '           End If
                '           If Trim(txt_mneutrophils.Text) = "" And Trim(txt_mlymphocytes.Text) = "" And Trim(txt_mmonocytes.Text) = "" And Trim(txt_meosinophils.Text) = "" And Trim(txt_mbasophills.Text) = "" Then
                '                CrystalReport1.Formulas(19) = "cbcm = '" & "" & "'"
                '           Else
                '                CrystalReport1.Formulas(19) = "cbcm = '" & "123" & "'"
                '           End If
                '           If Trim(TXT_HCT.Text) = "" And Trim(TXT_MCH.Text) = "" And Trim(TXT_MCV.Text) = "" And Trim(TXT_MCHC.Text) = "" And Trim(TXT_MPV.Text) = "" And Trim(TXT_RDW.Text) = "" And Trim(TXT_ABNORMAL.Text) = "" Then
                '                CrystalReport1.Formulas(20) = "rbci = '" & "" & "'"
                '           Else
                '                CrystalReport1.Formulas(20) = "rbci = '" & "123" & "'"
                '           End If
                '           If Trim(txt_neutrophils.Text) = "" And Trim(txt_lymphocytes.Text) = "" And Trim(txt_monocytes.Text) = "" And Trim(txt_eosinophils.Text) = "" And Trim(txt_basophils.Text) = "" And Trim(txt_immaturegranulocytes.Text) = "" Then
                '                CrystalReport1.Formulas(21) = "dlc = '" & "" & "'"
                '           Else
                '                CrystalReport1.Formulas(21) = "dlc = '" & "123" & "'"
                '           End If
                '           If Trim(txt_hb.Text) = "" And Trim(txt_rbc.Text) = "" And Trim(TXT_TWBC.Text) = "" And Trim(txt_platelets.Text) = "" Then
                '                CrystalReport1.Formulas(22) = "cbcc = '" & "" & "'"
                '           Else
                '                CrystalReport1.Formulas(22) = "cbcc = '" & "123" & "'"
                '           End If
                '           If Trim(txt_comments.Text) <> "" Then
                '                CrystalReport1.Formulas(23) = "nocomments='" & "" & "'"
                '                CrystalReport1.Formulas(24) = "comments='" & "COMMENTS:" & "'"
                '           Else
                '                CrystalReport1.Formulas(23) = "nocomments='" & "YES" & "'"
                '                CrystalReport1.Formulas(24) = "comments='" & "" & "'"
                '           End If
                '           '' PRINT REPORT
                '           CrystalReport1.WindowState = crptMaximized
                '           CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                '           CrystalReport1.Action = 1
                ''           If REP1 = True And REP2 = True Then
                ''              MsgBox "There are two reports print this one first"
                ''           End If
            End If
         End If
         End If '09/03/11
  
        'kanchan
        If Trim(txt_modcode.Text) = "HL" Then
'            str = ipadd
'            CrystalReport5.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'            CrystalReport5.ReportFileName = "C:\DISREPORTSEC\HEMATOstaff.RPT"
'
'            CrystalReport5.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'            CrystalReport5.Formulas(1) = "PROVDIA = '" & txt_procedure & "'"
'            CrystalReport5.WindowState = crptMaximized
'
'            CrystalReport5.Destination = crptToWindow
'            CrystalReport5.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'            CrystalReport5.Action = 1
            
            '***** Shraddha 06/05/2011 *****
            str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
            CrystalReport5.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
            
            str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
            
            CrystalReport5.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
            
            CrystalReport5.ReportFileName = "C:\DISREPORTSECLF\HEMATOstaff.RPT"

            CrystalReport5.Formulas(0) = "REQNO='" & Trim(cmb_reqno.Text) & "'"
               CrystalReport5.Formulas(5) = "PROVDIA = '" & txt_procedure & "'"
            CrystalReport5.Formulas(17) = "signature='" & Trim(cmb_finalby.Text) & "'"
               CrystalReport5.Formulas(8) = "USRID = '" & str & "'"
               
                CrystalReport5.WindowState = crptMaximized
                CrystalReport5.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport5.Action = 1
            '************ 06/05/2011 ******************
            
            
            ' End If
        End If
        '''  NEW
        '''*** FOR COAGULATION PROFILE STAFF
        If Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
            CrystalReport3.Formulas(0) = "REQNO='" & cmb_reqno.Text & "'"
            str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
            CrystalReport3.Formulas(1) = "USRID='" & str & "'"
            CrystalReport3.Formulas(2) = "PROFIN='" & "HAEMATO PATHOLOGY  REPORT" & "'"
            If Trim(txt_comments.Text) <> "" Then
                CrystalReport3.Formulas(3) = "nocomments='" & "" & "'"
                CrystalReport3.Formulas(4) = "COMMENTS='" & "COMMENTS:" & "'"
            Else
                CrystalReport3.Formulas(3) = "nocomments='" & "YES" & "'"
                CrystalReport3.Formulas(4) = "COMMENTS='" & "" & "'"
            End If
            If Trim(txt_impression.Text) <> "" Then
                CrystalReport3.Formulas(5) = "noimpress='" & "" & "'"
                CrystalReport3.Formulas(6) = "impress='" & "IMPRESSION:" & "'"
            Else
                CrystalReport3.Formulas(5) = "noimpress='" & "YES" & "'"
                CrystalReport3.Formulas(6) = "impress='" & "" & "'"
            End If
            If Trim(TXT_FDPDMIER.Text) <> "" Then
                CrystalReport3.Formulas(7) = "D-DIMER='" & "D-Dimer test is performed manually by Semi-Quantitative Latex Agglutination Technique." & "'"
            End If
            If Trim(TXT_TT.Text) <> "" Then
                CrystalReport3.Formulas(8) = "TT='" & "TT test is performed manually." & "'"
            End If
            'Dim STR As String
            ' STR = ""
            '            If Trim(TXT_PT(0).Text) <> "" Then
            '                STR = "PT"
            '
            '            End If
            '            If Trim(TXT_PT(0).Text) <> "" And Trim(TXT_PTTK.Text) <> "" Then
            '                STR = STR & "/APTT"
            '            ElseIf Trim(TXT_PT(0).Text) = "" And Trim(TXT_PTTK.Text) <> "" Then
            '                STR = STR & "APTT"
            '            End If
               
            If Trim(TXT_PT(0).Text) <> "" And Trim(TXT_PTTK.Text) <> "" And Trim(TXT_FIBRINOGEN.Text) <> "" Then
                str = "PT/APTT/FIBRINOGEN"
            ElseIf Trim(TXT_PT(0).Text) = "" And Trim(TXT_PTTK.Text) <> "" And Trim(TXT_FIBRINOGEN.Text) <> "" Then
                str = "APTT/FIBRINOGEN"
            ElseIf Trim(TXT_PT(0).Text) = "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_FIBRINOGEN.Text) <> "" Then
                str = "FIBRINOGEN"
            ElseIf Trim(TXT_PT(0).Text) <> "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_FIBRINOGEN.Text) <> "" Then
                str = "PT/FIBRINOGEN"
            ElseIf Trim(TXT_PT(0).Text) <> "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_FIBRINOGEN.Text) = "" Then
                str = "PT"
            ElseIf Trim(TXT_PT(0).Text) = "" And Trim(TXT_PTTK.Text) <> "" And Trim(TXT_FIBRINOGEN.Text) = "" Then
                str = "APTT"
            ElseIf Trim(TXT_PT(0).Text) <> "" And Trim(TXT_PTTK.Text) <> "" And Trim(TXT_FIBRINOGEN.Text) = "" Then
                str = "PT/APTT"
            End If
            
'            If Option1.Value = True Then
'                str1 = "ACL 100"
'            ElseIf Option2.Value = True Then
'                str1 = "ACL 200"
'            ElseIf Option3.Value = True Then
'                str1 = "ELP 01"
'            Else
                str1 = cmb_machname.Text
'            End If
            
            If Trim(TXT_PT(0).Text) <> "" Or Trim(TXT_PTTK.Text) <> "" Or Trim(TXT_FIBRINOGEN.Text <> "") Then
                CrystalReport3.Formulas(9) = "PT='" & str & " tests are performed on ACL automated coagulometer " & str1 & ".'"
            End If
            
            CrystalReport3.SelectionFormula = "{RREQUISITN.REQNNO}='" & req_chglen(cmb_reqno) & "'"
            CrystalReport3.ReportFileName = "C:\DISREPORTSEC\DISCPDETSTAFF.rpt"
            CrystalReport3.WindowState = crptMaximized
            CrystalReport3.Connect = "ODBC;Database=DISLIB,PATOOLSLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
            CrystalReport3.Action = 1
  
        End If '//CG
        
        If Trim(SELDHEMCYTO_RS("CYTO")) = "YES" Or Trim(SELDHEMCYTO_RS("SMAR")) = "YES" Then
   
            str = ipadd '& " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
            CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
   
            If Trim(txt_modcode.Text) = "HL" Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShlN1ASTAFF.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
                '        If Trim(prepath) = "" Then
                '            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                '        Else
                '            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
                '        End If
                CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
                
                If Trim(REFDR_RS("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "PERIPHERAL BLOOD REPORT" & "'"
                    str1 = ""
                    str1 = "Entered By " & Trim(cmb_enterby.Text)
                    If Len(Trim(cmb_enterby.Text)) > 0 Then
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                    Else
                        CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
                End If
            Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "PERIPHERAL BLOOD REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
           
            If Trim(txt_lapscore.Text) = "" And Trim(TXT_MPO.Text) = "" And Trim(cmb_nse.Text) = "" And Trim(TXT_PAS.Text) = "" And Trim(TXT_NSE.Text) = "" And Trim(TXT_TRAP.Text) = "" Then
                'If CYTOCHEMFLAG = True Then
                CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
            Else
                CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
                'End If
            End If
            
            If Trim(txt_inr.Text) = "" And Trim(TXT_PT(0).Text) = "" And Trim(txt_ptctl.Text) = "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_TT.Text) = "" And Trim(TXT_FIBRINOGEN.Text) = "" And Trim(TXT_FDPDMIER.Text) = "" And Trim(TXT_OTHERS.Text) = "" Then
                CrystalReport1.Formulas(11) = "CP = '" & "" & "'"
            Else
                CrystalReport1.Formulas(11) = "CP = '" & "ABC" & "'"
            End If

            If Trim(TXT_NORMOBLASTS.Text) = "" And Trim(TXT_OTHERCELLSRBC.Text) = "" And Trim(txt_rbcmorphology.Text) = "" And Trim(txt_reticcount.Text) = "" And Trim(txt_malarialparasite.Text) = "" And Trim(txt_esr.Text) = "" Then '18/11/08
                CrystalReport1.Formulas(15) = "RBC = '" & "" & "'"
            Else
                CrystalReport1.Formulas(15) = "RBC = '" & "ABC" & "'"
            End If
            'If Trim(TXT_MORPHOLOGY.Text) = "" And Trim(TXT_ADEQUACY.Text) = "" Then
            If Trim(CMB_MORPHO.Text) = "" And Trim(TXT_ADEQUACY.Text) = "" Then
                CrystalReport1.Formulas(16) = "PLATELETS = '" & "" & "'"
            Else
                CrystalReport1.Formulas(16) = "PLATELETS = '" & "ABC" & "'"
            End If
            If Trim(txt_prolymphocytes.Text) = "" And Trim(txt_Sezarycells.Text) = "" And Trim(txt_hairycells.Text) = "" And Trim(txt_othercells.Text) = "" Then
                CrystalReport1.Formulas(17) = "prolyn = '" & "" & "'"
            Else
                CrystalReport1.Formulas(17) = "prolyn = '" & "123" & "'"
            End If
            If Trim(txt_blasts.Text) = "" And Trim(TXT_PROMYELOCYTESWBC.Text) = "" And Trim(txt_myelocytes.Text) = "" And Trim(txt_metamyelocytes.Text) = "" And Trim(txt_plasmacells.Text) = "" And Trim(txt_alymphocytes.Text) = "" Then
                CrystalReport1.Formulas(18) = "blast = '" & "" & "'"
            Else
                CrystalReport1.Formulas(18) = "blast = '" & "123" & "'"
            End If
            If Trim(txt_mneutrophils.Text) = "" And Trim(txt_mlymphocytes.Text) = "" And Trim(txt_mmonocytes.Text) = "" And Trim(txt_meosinophils.Text) = "" And Trim(txt_mbasophills.Text) = "" Then
                CrystalReport1.Formulas(19) = "cbcm = '" & "" & "'"
            Else
                CrystalReport1.Formulas(19) = "cbcm = '" & "123" & "'"
            End If
            If Trim(TXT_HCT.Text) = "" And Trim(TXT_MCH.Text) = "" And Trim(TXT_MCV.Text) = "" And Trim(TXT_MCHC.Text) = "" And Trim(TXT_MPV.Text) = "" And Trim(TXT_RDW.Text) = "" And Trim(TXT_ABNORMAL.Text) = "" Then
                CrystalReport1.Formulas(20) = "rbci = '" & "" & "'"
            Else
                CrystalReport1.Formulas(20) = "rbci = '" & "123" & "'"
            End If
            If Trim(txt_neutrophils.Text) = "" And Trim(txt_lymphocytes.Text) = "" And Trim(txt_monocytes.Text) = "" And Trim(txt_eosinophils.Text) = "" And Trim(txt_basophils.Text) = "" And Trim(txt_immaturegranulocytes.Text) = "" Then
                CrystalReport1.Formulas(21) = "dlc = '" & "" & "'"
            Else
                CrystalReport1.Formulas(21) = "dlc = '" & "123" & "'"
            End If
            If Trim(TXT_HB.Text) = "" And Trim(TXT_RBC.Text) = "" And Trim(TXT_TWBC.Text) = "" And Trim(txt_platelets.Text) = "" Then
                CrystalReport1.Formulas(22) = "cbcc = '" & "" & "'"
            Else
                CrystalReport1.Formulas(22) = "cbcc = '" & "123" & "'"
            End If
            '           If Trim(SELDHEMCYTO_RS("smar")) = "YES" Then
            '             CrystalReport1.Formulas(23) = "mar = '" & "" & "'"
            '           Else
            '             CrystalReport1.Formulas(23) = "mar = '" & "123" & "'"
            '           End If
           
            ' PRINT REPPORT
           
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
            CrystalReport1.Action = 1
        End If
    End If  'CYTO/SURFACE MARKER
     
    ' NEW CODES FOR STAFF REPORTS
    '  OLD CODE FOR STAF REPORT FOR HL BEGINS
    '        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & Trim(VAR_REQNNO) & "'"
    '        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShlSTAFF.RPT"
    '            CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
    ''        If Trim(prepath) = "" Then
    ''            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
    ''        Else
    ''            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
    ''        End If
    '        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
    '        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
    '            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
    '            STR1 = ""
    '            STR1 = "Entered By " & Trim(cmb_enterby.Text)
    '            If Len(Trim(cmb_enterby.Text)) > 0 Then
    '                CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
    '            Else
    '                 CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
    '            End If
    '            Else
    '            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
    '            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
    '            End If
    '        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
    '
    '        If Trim(txt_lapscore.Text) = "" And Trim(TXT_MPO.Text) = "" And Trim(cmb_nse.Text) = "" And Trim(TXT_PAS.Text) = "" And Trim(TXT_NSE.Text) = "" And Trim(TXT_TRAP.Text) = "" Then
    '               'If CYTOCHEMFLAG = True Then
    '               CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
    '        Else
    '              CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
    '               'End If
    '        End If
    '        If Trim(txt_inr.Text) = "" And Trim(TXT_PT(0).Text) = "" And Trim(txt_ptctl.Text) = "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_TT.Text) = "" And Trim(TXT_FIBRINOGEN.Text) = "" And Trim(TXT_FDPDMIER.Text) = "" And Trim(txt_others.Text) = "" Then
    '            CrystalReport1.Formulas(11) = "CP = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(11) = "CP = '" & "ABC" & "'"
    '        End If
    '
    '        If Trim(TXT_NORMOBLASTS.Text) = "" And Trim(TXT_OTHERCELLSRBC.Text) = "" And Trim(txt_rbcmorphology.Text) = "" And Trim(txt_reticcount.Text) = "" And Trim(txt_malarialparasite.Text) = "" And Trim(TXT_ESRRBC.Text) = "" Then
    '             CrystalReport1.Formulas(15) = "RBC = '" & "" & "'"
    '        Else
    '               CrystalReport1.Formulas(15) = "RBC = '" & "ABC" & "'"
    '        End If
    '        If Trim(TXT_MORPHOLOGY.Text) = "" And Trim(TXT_ADEQUACY.Text) = "" Then
    '            CrystalReport1.Formulas(16) = "PLATELETS = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(16) = "PLATELETS = '" & "ABC" & "'"
    '        End If
    '        If Trim(txt_prolymphocytes.Text) = "" And Trim(txt_Sezarycells.Text) = "" And Trim(txt_hairycells.Text) = "" And Trim(txt_othercells.Text) = "" Then
    '            CrystalReport1.Formulas(17) = "prolyn = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(17) = "prolyn = '" & "123" & "'"
    '        End If
    '        If Trim(txt_blasts.Text) = "" And Trim(TXT_PROMYELOCYTESWBC.Text) = "" And Trim(txt_myelocytes.Text) = "" And Trim(txt_metamyelocytes.Text) = "" And Trim(txt_plasmacells.Text) = "" And Trim(txt_alymphocytes.Text) = "" Then
    '            CrystalReport1.Formulas(18) = "blast = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(18) = "blast = '" & "123" & "'"
    '        End If
    '        If Trim(txt_mneutrophils.Text) = "" And Trim(txt_mlymphocytes.Text) = "" And Trim(txt_mmonocytes.Text) = "" And Trim(txt_meosinophils.Text) = "" And Trim(txt_mbasophills.Text) = "" Then
    '            CrystalReport1.Formulas(19) = "cbcm = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(19) = "cbcm = '" & "123" & "'"
    '        End If
    '        If Trim(TXT_HCT.Text) = "" And Trim(TXT_MCH.Text) = "" And Trim(TXT_MCV.Text) = "" And Trim(TXT_MCHC.Text) = "" And Trim(TXT_MPV.Text) = "" And Trim(TXT_RDW.Text) = "" And Trim(TXT_ABNORMAL.Text) = "" Then
    '            CrystalReport1.Formulas(20) = "rbci = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(20) = "rbci = '" & "123" & "'"
    '        End If
    '        If Trim(txt_neutrophils.Text) = "" And Trim(txt_lymphocytes.Text) = "" And Trim(txt_monocytes.Text) = "" And Trim(txt_eosinophils.Text) = "" And Trim(txt_basophils.Text) = "" And Trim(txt_immaturegranulocytes.Text) = "" Then
    '            CrystalReport1.Formulas(21) = "dlc = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(21) = "dlc = '" & "123" & "'"
    '        End If
    '        If Trim(txt_hb.Text) = "" And Trim(TXT_RBC.Text) = "" And Trim(TXT_TWBC.Text) = "" And Trim(txt_platelets.Text) = "" Then
    '            CrystalReport1.Formulas(22) = "cbcc = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(22) = "cbcc = '" & "123" & "'"
    '        End If
    '  OLD CODE FOR STAF REPORT FOR HL ENDS
End If
         
If REP1 = True And REP2 = True Then
              'MsgBox "There are two reports print each one of them"
End If
'End If
'********** Shraddha Pandit 22/09/09 **********
txt_NEW = NEWFLAG
txt_VIEW = viewflag
txt_SAVE = flagsave
'**********************************************

'***** STAFF
'CrystalReport1.WindowState = crptMaximized
'CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'CrystalReport1.Action = 1
Exit Sub
prtmsg: MsgBox ERROR(Err)

Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub

Private Sub cmd_save_Click()
If Left(cmb_reqno.Text, 3) = "UZZ" And SSTab1.Tab = 2 Then
    If CMB_range.Text = "select" Or CMB_range.Text = "" Then
MsgBox ("Please Select D-Dimer range value")
Exit Sub
End If
End If
Dim pcount As Integer
pcount = 0
Dim beforst As String, afterst As String
Dim K As Long
Dim DCBCFLAG As Boolean
Dim tot As Integer
       
tot = 0

'SSTab1.Caption

If SSTab1.TabEnabled(1) = True Then
    If Len(Trim(txt_blasts.Text)) = 0 Then
        txt_blasts.Text = 0
    End If
    If Len(Trim(TXT_PROMYELOCYTESWBC.Text)) = 0 Then
        TXT_PROMYELOCYTESWBC.Text = 0
    End If
    If Len(Trim(txt_promonocyte.Text)) = 0 Then
        txt_promonocyte.Text = 0
    End If
    If Len(Trim(txt_myelocytes.Text)) = 0 Then
        txt_myelocytes.Text = 0
    End If
    If Len(Trim(txt_metamyelocytes.Text)) = 0 Then
        txt_metamyelocytes.Text = 0
    End If
    If Len(Trim(txt_mneutrophils.Text)) = 0 Then
        txt_mneutrophils.Text = 0
    End If
    If Len(Trim(txt_mlymphocytes.Text)) = 0 Then
        txt_mlymphocytes.Text = 0
    End If
    '******** 16/10/08 Shraddha Pandit *******
    If Len(Trim(txt_prolympho.Text)) = 0 Then
        txt_prolympho.Text = 0
    End If
    '*****************************************
    If Len(Trim(txt_mmonocytes.Text)) = 0 Then
        txt_mmonocytes.Text = 0
    End If
    If Len(Trim(txt_meosinophils.Text)) = 0 Then
        txt_meosinophils.Text = 0
    End If
    If Len(Trim(txt_mbasophills.Text)) = 0 Then
        txt_mbasophills.Text = 0
    End If
    If Len(Trim(txt_plasmacells.Text)) = 0 Then
        txt_plasmacells = 0
    End If
    If Len(Trim(txt_alymphocytes.Text)) = 0 Then
        txt_alymphocytes = 0
    End If
    If Len(Trim(txt_Sezarycells.Text)) = 0 Then
        txt_Sezarycells.Text = 0
    End If
    If Len(Trim(txt_hairycells.Text)) = 0 Then
        txt_hairycells.Text = 0
    End If
    If Len(Trim(txt_othercells.Text)) = 0 Then
        txt_othercells.Text = 0
    End If
            
    'tot = CInt(Trim(txt_blasts.Text)) + CInt(Trim(TXT_PROMYELOCYTESWBC.Text)) + CInt(Trim(txt_promonocyte.Text)) + CInt(Trim(txt_myelocytes.Text)) + CInt(Trim(txt_metamyelocytes.Text)) + CInt(Trim(txt_mneutrophils.Text))
    tot = tot + CInt(Trim(txt_mlymphocytes.Text)) + CInt(Trim(txt_mmonocytes.Text)) + CInt(Trim(txt_meosinophils.Text)) + CInt(Trim(txt_mbasophills.Text)) + CInt(Trim(txt_plasmacells.Text))
    tot = tot + CInt(Trim(txt_alymphocytes.Text)) + CInt(Trim(txt_Sezarycells.Text)) + CInt(Trim(txt_hairycells.Text)) + CInt(Trim(txt_othercells.Text))

'    If tot <> 100 Then
'        If tot < 100 Then
'            MsgBox ("Total cell count can not be less than 100..."), vbOKOnly
'        ElseIf tot > 100 Then
'            MsgBox ("Total cell count can not be greater than 100..."), vbOKOnly
'        End If
'        '   Exit Sub
'    End If
End If
       
If DCBCFLAG = False Then
    SELDCBC_QR(0) = req_chglen(cmb_reqno.Text)
    Set SELDCBC_RS = SELDCBC_QR.OpenResultset(1, 3)
    DCBCFLAG = True
Else
    SELDCBC_QR(0) = req_chglen(cmb_reqno.Text)
    SELDCBC_RS.MoreResults
    SELDCBC_RS.Requery
End If
      
If SELDCBC_RS.RowCount > 0 Then
    flagsave = 2 'update
Else
    flagsave = 1 'insert
End If

'If Len(Trim(txt_impression.Text)) = 0 Then
'    MsgBox "Impression field is Empty.", vbCritical
'    Exit Sub
'End If
  
'***** 13/08/2008 NO NEED OF ENTERED BY ******
'If Len(Trim(cmb_enterby.Text)) = 0 Then
'    MsgBox "Enter by field is Empty.", vbCritical
'    Exit Sub
'End If

If Len(Trim(txt_impression.Text)) > 1000 Then
    MsgBox "Max. of Impression field length is 1000 characters", vbCritical
    Exit Sub
End If

If Len(Trim(cmb_finalby.Text)) = 0 Then
    MsgBox "To be finalised  field is Empty.", vbCritical
    Exit Sub
End If
If Left(cmb_reqno.Text, 3) = "UZZ" Then
If Len(Trim(regcbo.Text)) = 0 Then
    MsgBox "Registrar  field is Empty.", vbCritical
    Exit Sub
End If
End If
If Len(Trim(txt_comments.Text)) > 2000 Then
    MsgBox "Max. of Comments field length is 2000 characters", vbCritical
    Exit Sub
End If
    
' If SSTab1.Tab = 3 Or SSTab1.Tab = 4 Then
'    If TXT_DNAINDEX.Text = "" Or TXT_DNACLOIDY.Text = "" Then
'    MsgBox "Please Enter values in DNA Index and DNA Cloidy both", vbExclamation
'    TXT_DNAINDEX.SetFocus
'    Exit Sub
'    End If
' End If
    
    
assign

btnenabl ("SAVE")

'********** Shraddha Pandit 22/09/09 **********
  txt_NEW = NEWFLAG
  txt_VIEW = viewflag
  txt_SAVE = flagsave
'************************************
If Left(cmb_reqno.Text, 3) = "UZZ" And SSTab1.Tab = 2 Then

Dim STR6, STR7, DDIMER_RANGE
Dim range_n_qr1 As rdoQuery
Dim range_n_rs1 As rdoResultset
      Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
                DDIMER_RANGE = SELCPRNG_RS("RANGE")
                
 ST7 = "UPDATE dislib.DHEMCpdet SET NRANGE='" & DDIMER_RANGE & "',OBSERTYPE='" & CMB_range.Text & "' where reqno='" & req_chglen(cmb_reqno.Text) & "'   and PROCCODE='D-DIMER' "
STR6 = "select * from dislib.DHEMCpdet where reqno='" & req_chglen(cmb_reqno.Text) & "' and PROCCODE='D-DIMER'"
       Set range_n_qr = rdb.CreateQuery("", STR6)

Set range_n_rs = range_n_qr.OpenResultset(1, 1)

If range_n_rs.RowCount > 0 Then
 Set range_n_qr1 = rdb.CreateQuery("REPW0", ST7)

Set range_n_rs1 = range_n_qr1.OpenResultset(1, 1)

End If
End If
'cmd_print.Enabled = True
'cmd_cancel.Enabled = True
'cmd_commit.Enabled = True

End Sub

Public Sub cmd_view_Click()
'On Error GoTo MSG:
txt_1.Text = ""
btnenabl ("VIEW")
FLDENABL ("DIS")
field_clear
'TXT_CLEAR
Command1.Visible = False
'***** Shraddha 30/03/09 ********
ADDENDUMFLAG = False
'********************************
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
If sel_deptrs.RowCount > 0 Then
    sel_deptrs.MoveFirst
    While Not sel_deptrs.EOF
        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
        If (Trim(USRDEPT) = "FED" Or Trim(USRDEPT) = "FEC") And Trim(sel_deptrs("dept_code")) = "FED" Then
    cmb_dept.Text = Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
          End If
        sel_deptrs.MoveNext
        Wend
   End If
 
'         If seldissecentrflag1 = False Then
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
 
   
cmb_dept.Enabled = True
ADDENDUMFALG = False
viewflag = True
Opt_patient.Value = True
cmb_dept.SetFocus
cmb_modality.Enabled = True
MaskEdBox1.Enabled = True
cmb_reqno.Enabled = True
cmd_print.Enabled = True

'********** Shraddha Pandit 22/09/09 **********
  txt_NEW = NEWFLAG
  txt_VIEW = viewflag
  txt_SAVE = flagsave
'**********************************************

Exit Sub


MSG:
MsgBox Err.Description
End Sub

Private Sub Command1_Click()
TXT_HB.Text = ""
TXT_RBC.Text = ""
TXT_TWBC.Text = ""
txt_platelets.Text = ""
diffcountflag = True
End Sub

Private Sub Command2_Click()
If Len(Trim(MaskEdBox1.Text)) > 0 Then
    VAR1 = MaskEdBox1.FormattedText & "~" & Replace(Trim(cmb_reqno.Text), "/", "") & "~" & usr_name
    Dim strLink As String
    'strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
    strLink = "http://intranet.tmc.gov.in/imunorpt/dis_all_1.ASPx?CASE_NO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If

End Sub

Private Sub Form_Load()
'On Error GoTo MSG:

 
NEWFLAG = False
viewflag = False
FLDENABL ("DIS")
ENABL ("DIS")
DATETIME
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
If HAEMATOLOGYFLAG = False Then
    HAEMATOLOGYFLAG = True
  

'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1BIO", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enterby.Text = (LOGINUSRNAME)
    cmb_enterby.AddItem LOGINUSRNAME
    cmb_enterby.Text = cmb_enterby.List(0)
    cmb_enterby.Enabled = False

Else
    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'***** Shraddha 01/03/11 *****
'txt_cbc_comment.Text = "Manually verified values will be sent as a separate report."
'Commented again on 09/03/11
'*****************************



'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
'Commented and added by Shweta on 20-06-2013 as told by Mrs.Kalpana madam
'str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? ORDER BY USERNAME"
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? and (MODMO IN ('5','6','7')) ORDER BY USERNAME"
'Commented and added by Shraddha on 21/12/09 as told by Mrs.Surekha
'str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE=? and USRTYPE='8' ORDER BY USERNAME"
'*********************
Set REPWTRN_QR = rdb.CreateQuery("REPWTRNSTR1", str)
If Trim(USRDEPT) = "ALL" Then
    REPWTRN_QR(0) = "UZZ"
Else
    REPWTRN_QR(0) = Trim(USRDEPT)
End If
Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)

If REPWTRNRS.RowCount > 0 Then
    'cmb_enteryby.clear
    cmb_addendumby.clear
    'cmb_finalby.clear
    regcbo.clear
    'cmb_provby.clear
    REPWTRNRS.MoveFirst
    While Not REPWTRNRS.EOF
        If REPWTRNRS("USRTYPE") > 5 And (Trim(REPWTRNRS("MODMO")) = "5" Or Trim(REPWTRNRS("MODMO")) = "7") Then 'CHANGES MADE FOR CBC TEST
            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
            
        ElseIf REPWTRNRS("USRTYPE") = 5 And (Trim(REPWTRNRS("MODMO")) = "6" Or Trim(REPWTRNRS("MODMO")) = "7") Then
            regcbo.AddItem Trim(REPWTRNRS("USERNAME"))
            'cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
        'ElseIf REPWTRNRS("USRTYPE") = 4 And REPWTRNRS("USRTYPE") = 9 Then
            
        
        End If
        REPWTRNRS.MoveNext
    Wend
    'cmb_finalby.Text = cmb_finalby.List(0)
Else
    MsgBox "The Security Details of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11BIO", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

''**********SECURITY****************************

btnenabl ("INI")

''*******ADDENDUM***********
    str = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=? "
    Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR1111XX", str)
'
    SELDADDENDUM_QR(0) = "FSS"
    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'
'    STR = "select * from dislib.dissec where ADDENDUMBY='YES'AND MODALITY='HL' "
'    Set seldissecentr11_QR = rdb.CreateQuery("HLseldissecentr1SS_qr1BM", STR)
'    Set seldissecentr11_RS = seldissecentr11_QR.OpenResultset(1, 3)
'    If Not seldissecentr11_RS.EOF Then
'        Do While Not seldissecentr11_RS.EOF
'            cmb_addendumby.AddItem Trim(seldissecentr11_RS("CCNAME"))
'            seldissecentr11_RS.MoveNext
'        Loop
'    End If

''*******ADDENDUM***********
  
  
   ''************FOR PREPATHNO & REQNO COMBOBOX REPORT ****
''DCBCREPORT FOR HL
    str = "SELECT * FROM DISLIB.DCBCREP1 WHERE REQNO=?"
    Set SELDCBCREPORT_QR = rdb.CreateQuery("SELDCBCREPORT_CBCSTRSDDED", str)
    
    ''DCBCREPORT FOR HL
    str = "SELECT * FROM DISLIB.DBMAREPORT WHERE REQNO=?"
    Set SELDBMAREP_QR = rdb.CreateQuery("SELDBMAREPORT_CBCSTRCVGF", str)
    
    'STR = "SELECT * FROM RISLIB.RREQUISITN LEFT OUTER JOIN DISLIB.DREQUISITN ON RISLIB.RREQUISITN.REQNNO=DISLIB.DREQUISITN.REQNO WHERE RISLIB.RREQUISITN.CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')" 'ORDER BY RREQUISITN.REQNDT DESC"
    str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')"
    Set seldrequisitn_qr10 = rdb.CreateQuery("seldrequisitn11CBC_1qr11XX0", str)
         
    str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND REQNSTAT='F' AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'LZZCR%' OR REQNNO LIKE 'LZZCT%' OR REQNNO LIKE 'LZZMM%' OR REQNNO LIKE 'LZZMR%' OR REQNNO LIKE 'LZZUS%')"
    Set seldrequisitn_qr11 = rdb.CreateQuery("seldrequisitn11CBC_1qrXXREP", str)
     
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
    Set SELRREQNOQR = rdb.CreateQuery("SELRREQNOQR1SPREQXXCBC", str)
     
    'STR = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=?"
    'Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR11SPREQXXCBC", STR)
        
    
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN1SPREQ_QR1XXCBC", str)
    
    str = "SELECT REQNO,FSPREPORT,IHREPORT,IMPRESSION,GROSSDESC,GROSSEDBY,SPREPORT,REPORTIND,FNAC FROM DISLIB.DFREPORT WHERE REQNO=?"
    Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT_1QR1sprepXXCBC", str)
    
    str = "SELECT * FROM PATOOLSLIB.TA1012 WHERE CASE_NO=?"
    Set selta1012_qr = rdb.CreateQuery("SELTA1012_QR11klCBC", str)
    selta1012_qr(0) = "TT/00333"
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    
    str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
    Set SELMA301QR = rdb.CreateQuery("SELMA301QR1QSWCBC", str)
    
    SELMA301QR(0) = "F22"
    Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
    
    str = "SELECT REQNNO,REQNDT,REQNSTAT,CASENO,PREPORTDT,PATHNO FROM DISLIB.DSPLF WHERE CASENO=? AND (REQNSTAT='T' OR REQNSTAT='F') AND (REQNNO LIKE 'FZZCG%' OR REQNNO LIKE 'FZZSP%' OR REQNNO LIKE 'FZZCN%' OR  REQNNO LIKE 'FZZFA%' OR REQNNO LIKE 'UZZ%') ORDER BY REQNNO"
    Set PREPATH_QR = rdb.CreateQuery("PREPATH_STRcbc1", str)
    
    '******* Shraddha 30/03/09
    str = "UPDATE RISLIB.RTIMELOG SET ADRREPDATE=?, ADRREPTIME=? WHERE REQNNO=?"
    Set UPDTADNDMTIME_QR = rdb.CreateQuery("UPDTADNDMTIME_QR", str)
    
    '*********************************
    
    '***** Shraddha 09/03/10 *****

    'str = "SELECT * FROM DISLIB.HINTERFACE WHERE REQNO=? and MACHINENO='A2120'"
    str = "SELECT * FROM DISLIB.HINTERFACE WHERE REQNO=? and MACHINENO LIKE 'A212%'"
    Set SEL_DHINTER_QR = rdb.CreateQuery("", str)
    
    '***** Shraddha 09/03/11 *****
    str = "SELECT * FROM DISLIB.REPORTHLF1 WHERE reqno=? and OBSER_TYPE=?"
    Set sel_dhemcbc_qr = rdb.CreateQuery("", str)

    str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=?"
    Set SELSTAFFINFO_QR = rdb.CreateQuery("", str)

    '**********************************
    
    
    ''************FOR PREPATHNO & REQNO COMBOBOX REPORT ****
   
'*********SECURITY**********************************

'STR = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
'Set USRNAME_QR = rdb.CreateQuery("USRNAME1CBC", STR)
'USRNAME_QR(0) = UID
'Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
'If USRNAME_RS.RowCount > 0 Then
'    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
'    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
'    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
'    cmb_enteryby.AddItem LOGINUSRNAME
'Else
'    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
'    cmd_cancel_Click
'    Exit Sub
'End If
'
'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
'Set REPWTRNRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USERNAME", 1, 1)
'If REPWTRNRS.RowCount > 0 Then
'    'cmb_enteryby.clear
'    cmb_addendumby.clear
'    cmb_finalby.clear
'    cmb_provby.clear
'    REPWTRNRS.MoveFirst
'    While Not REPWTRNRS.EOF
'        If REPWTRNRS("USRTYPE") > 1 Then
'            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
'            'cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
'            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
'            cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
'        End If
'        REPWTRNRS.MoveNext
'    Wend
'Else
'    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
'    cmd_cancel_Click
'    Exit Sub
'End If
'STR = "select * from DISLIB.Dislnkusr where username = ?"
'Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11CBC", STR)
'CURUSRDET_QR(0) = "ABC"
'Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)
'
'STR = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE=?"
'Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1CBC", STR)
'SELFINALBY_QR(0) = 6
'Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
'If Not SELFINALBY_RS.EOF Then
'    Do While Not SELFINALBY_RS.EOF
'       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
'       SELFINALBY_RS.MoveNext
'    Loop
'End If


'*********SECURITY**********************************
''for dept selection
str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ' OR DEPT_CODE='FED')"
Set sel_deptqr = rdb.CreateQuery("sel_deptqr1hE", str)
'' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel_moadlityqr1HE", str)

'    'for retriving reqnos
Set req_qr = rdb.CreateQuery("reqBM1", "select * from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E'AND MODALITYCD=? ORDER BY REQNDT desc ")

'
Set SELTA101_QR = rdb.CreateQuery("selta101_qr11BM", "select * from PATOOLSLIB.ta101 where case_no =?")

'    ' FOR REFEFING DOCTOR
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("REFDR_QR1HE", str)

'    ' FOR PATH NO
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
    Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN_QR1HE", str)
''
'    ' for bedno, ward retriveing
    str = "select * from PATOOLSLIB.ma104 where case_no =?"
    Set selma104_qr = rdb.CreateQuery("selma104_qr11H", str)

'
    str = "select * from patoolslib.ta1012 where case_no =?"
    Set selta1012_qr = rdb.CreateQuery("selta1012_qr1HE", str)

    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?) ORDER BY REQNDT DESC "
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR1HE", str)

   str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn_qr1HE", str)
'
   str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
   Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn1_qr1NHE", str)
'
   str = "INSERT INTO DISLIB.DADDENDUM(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
   Set INSDADDENDUM_QR = rdb.CreateQuery("INSDADDENDUMSS_QR1", str)
   
   str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT=?,COMMITFLAG=? WHERE REQNNO=?"
   Set UPDRREQ_QR = rdb.CreateQuery("UPDATERREQ", str)

'
'   STR = "SELECT  * FROM DISLIB.DADDENDUM WHERE REQNO=?"
'   Set SELDADDENDUM_QR = rdb.CreateQuery("selDADDENDUM_QR11", STR)
'   SELDADDENDUM_QR(0) = "FSS"
'   Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'
   str = "select * from dislib.drequisitn where caseno=?"
   Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_qr1HE", str)

    'commented and new query added by amarjeet on 05-11-2002 to select manjure's name in enter by
    str = "select * from dislib.dissec where enterby='YES'AND MODALITY='HL' "
 
   Set seldissecentr1_qr = rdb.CreateQuery("seldissecentr1_qr1HE", str)
   
   str = "select * from dislib.dissec where PROVby='YES'AND MODALITY='HL' OR MODALITY='BC'"
   Set seldissecprov1_qr = rdb.CreateQuery("seldissecPROV1_qr1HE", str)
'
   str = "select * from dislib.dissec where FINALby='YES'AND MODALITY='HL' OR MODALITY='BC'"
   Set seldissecfin1_qr = rdb.CreateQuery("seldissecFIN1_qr1HE", str)
    
    str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
    Set SELDREQSTAFFDQR = rdb.CreateQuery("SELDREQSTAFFDQR1HE", str)
    
  str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
  Set SELDMULPROCQR = rdb.CreateQuery("SELDMULPROCQRHE", str)
    
  'STR = "SELECT * FROM DISLIB.MA101 WHERE SER_CODE=?"
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
  Set selMA101QR = rdb.CreateQuery("SELMA101QRHE", str)
  
    'STR = "update INTERFACE.hinterface set iflag = '3' where reqno = ?"
    'Set updhinterface_qr = rdb.CreateQuery("updhinter", STR)
    
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_DET=?"
  Set SELDMA101QR = rdb.CreateQuery("SELDMA101QRHE", str)
  
  str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,freportby=?,freportdt=? WHERE REQNNO=?"
   Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn11_qr1HE", str)
  
  
  str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
  Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qrHE", str)
  
  str = "SELECT * FROM DISLIB.DREQUISITN WHERE CASENO=? AND REQNO LIKE '%EC%'"
  Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQRHE", str)
                                              '?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
  str = "INSERT INTO DISLIB.DCBCREP1 VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
  Set INSDBMAREPORT_QR = rdb.CreateQuery("INSDBMAREPORT_QRHE", str)
  
  ''''KANCHAN
  'str = "INSERT INTO DISLIB.DCBCREPA (REQNO,CRC,RPI,EMPCNT,MALERIA,FLD2) VALUES (?,?,?,?,?,?)"
  'Commented and added by Shraddha on 01/03/11
  'str = "INSERT INTO DISLIB.DCBCREPA (REQNO,CRC,RPI,EMPCNT,MALERIA,FLD2,FLD3) VALUES (?,?,?,?,?,?,?)"
  'Commented and added by SHWETA on 10-02-2014
  str = "INSERT INTO DISLIB.DCBCREPA (REQNO,CRC,RPI,EMPCNT,MALERIA,FLD2,FLD3,FLD7,FLD6) VALUES (?,?,?,?,?,?,?,?,?)"
  Set INSDCBCREPA_QR = rdb.CreateQuery("INSERTDCBCREPA", str)
  
  str = "SELECT * FROM DISLIB.DHEMCBCDET WHERE REQNO=?"
  Set VEWDHEMCBCDET_QR = rdb.CreateQuery("VIEWDHEMCBCDET", str)
  
  
  'str = "UPDATE DISLIB.DCBCREPA SET CRC=?,RPI=?,EMPCNT=?,MALERIA=?,FLD2=? WHERE REQNO=?"
  'Commented and added by Shraddha on 01/03/11 *****
 ' str = "UPDATE DISLIB.DCBCREPA SET CRC=?,RPI=?,EMPCNT=?,MALERIA=?,FLD2=?, FLD3=? WHERE REQNO=?"
  'Commented and added by SHWETA on 10-02-2014
  str = "UPDATE DISLIB.DCBCREPA SET CRC=?,RPI=?,EMPCNT=?,MALERIA=?,FLD2=?, FLD3=?,FLD7=?, FLD6=? WHERE REQNO=?"
  Set UPDDCBCREPA_QR = rdb.CreateQuery("UPDATEDCBCREPA", str)
  
  
  str = "SELECT * FROM DISLIB.DCBCREPA WHERE REQNO=?"
  Set SELDCBCREPA_QR = rdb.CreateQuery("SELECTDCBCREPA", str)
  
  str = "SELECT * FROM DISLIB.DCBCREP1 WHERE REQNO=?"
  Set SELDCBC_QR = rdb.CreateQuery("SELECTDCBCR", str)
  
  str = "INSERT INTO DISLIB.DHEMCBCDET VALUES(?,?,?,?,?,?,?,?,?,?)"
  Set INSDHEMCBCDET_QR = rdb.CreateQuery("INSDHEMCBCDET_QR", str)
  
  str = " UPDATE DISLIB.DHEMCBCDET SET OBSERVALUE=?, OBSER_TYPE=?, SOURCE=?, ENTRYBY=?, REMARKS=?, UOM=?, NRANGE=?, PVALUE=? WHERE REQNO=? AND PRO_CODE=?"
  Set UPDDHEMCBCDET_QR = rdb.CreateQuery("UPDDHEMCBCDET_QR", str)
  
  str = "SELECT * FROM INTERFACE.HINTERFACE WHERE REQNO=?"
  Set SELHINTFVAL_QR = rdb.CreateQuery("SELHINTFVAL_QR", str)
    
  str = "SELECT * FROM DISLIB.DHEMCBCDET WHERE REQNO=?"
  Set SELDCBCDET_QR = rdb.CreateQuery("SELDCBCDET_QR", str)
  
  str = "SELECT * FROM DISLIB.DHEMCBCDET WHERE REQNO=?"
  Set SELDHEMCBCDET_QR = rdb.CreateQuery("SELDHEMCBCDET_QR", str)
  
  str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
  Set UPDRTIMEC_QR = rdb.CreateQuery("QWEQWEQWEccy", str)

  str = "SELECT * FROM DISLIB.DHEMCBCRNG WHERE PROCODE=? AND AGE2=? AND SEX=? "
  Set SELDHEMCBCQR = rdb.CreateQuery("SELDHEMCBCQR", str)

  'Commented and added by 06/05/2011
  str = "UPDATE DISLIB.DCBCREP1 SET hb=?,rbc=?,tlc=?,platelets=?,neutrophil=?,lymphocyte=?,monocytes=?,eosinophil=?,basophils=?,immgranulo=?,hct=?,mch=?,mcv=?,mchc=?,mpv=?,rdw=?,abnormal=?,neutrophi1=?,lymphocyt1=?,monocyte1=?,eosinophi1=?,basophill1=?,blasts=?,myelocytes=?,metamyeloc=?,plasmacell=?,atlymphee=?,prolymphoc=?,sezarycell=?,hairycells=?,othercells=?,normboblas=?,othercellr=?,rbcmorpho=?,reticcount=?,mparasite=?,esr=?,PROMYELOCY=?,bt=?,pt=?,platelet1=?,pttk=?,tt=?,fibrinogen=?,fdpddimer=?,others=?,lap=?,mpo=?,sbb=?,psa=?,nse=?,trad=?,fld1=?,impression=?, PMORPHOLOG=?,apttctl=?,ttctl=?,lapctl=?,mpoval=?,comments=? WHERE REQNO=?"
  'str = " UPDATE DISLIB.DCBCREP1 SET hb=?,rbc=?,tlc=?,platelets=?,neutrophil=?,lymphocyte=?,monocytes=?,eosinophil=?,basophils=?,immgranulo=?,hct=?,mch=?,mcv=?,mchc=?,mpv=?,rdw=?,abnormal=?,neutrophi1=?,lymphocyt1=?,monocyte1=?,eosinophi1=?,basophill1=?,blasts=?,promyelocy=?,myelocytes=?,metamyeloc=?,plasmacell=?,atlymphee=?,prolymphoc=?,sezarycell=?,hairycells=?,othercells=?,normboblas=?,othercellr=?,rbcmorpho=?,reticcount=?,mparasite=?,esr=?,PLTMORPHO=?,bt=?,pt=?,platelet1=?,pttk=?,tt=?,fibrinogen=?,fdpddimer=?,others=?,lap=?,mpo=?,sbb=?,psa=?,nse=?,trad=?,fld1=?,impression=?, PMORPHOLOG=?,apttctl=?,ttctl=?,lapctl=?,mpoval=?,comments=? WHERE REQNO=?"
  'STR = " UPDATE DISLIB.DCBCREP1 SET hb=?,rbc=?,tlc=?,platelets=?,neutrophil=?,lymphocyte=?,monocytes=?,eosinophil=?,basophils=?,immgranulo=?,hct=?,mch=?,mcv=?,mchc=?,mpv=?,rdw=?,abnormal=?,neutrophi1=?,lymphocyt1=?,monocyte1=?,eosinophi1=?,basophill1=?,blasts=?,promyelocy=?,myelocytes=?,metamyeloc=?,plasmacell=?,atlymphee=?,prolymphoc=?,sezarycell=?,hairycells=?,othercells=?,normboblas=?,othercellr=?,rbcmorpho=?,reticcount=?,mparasite=?,PLTMORPHO=?,bt=?,pt=?,platelet1=?,pttk=?,tt=?,fibrinogen=?,fdpddimer=?,others=?,lap=?,mpo=?,sbb=?,psa=?,nse=?,trad=?,fld1=?,impression=?, PMORPHOLOG=?,apttctl=?,ttctl=?,lapctl=?,mpoval=?,comments=? WHERE REQNO=?"
  Set UPDDBMAREPORT_QR = rdb.CreateQuery("UPDDBMAREPORT_QRHE", str)
  
  str = "SELECT * FROM DISLIB.DCBCREP1 WHERE REQNO=?"
  Set SELDBMAREPORT_QR = rdb.CreateQuery("SELDBMAREPORT_QRHE", str)
      
  str = "INSERT INTO DISLIB.DCBCSUBREP (REQNO,TESTNO,TESTDATE,HB,RBC,TLC,PLATELETS) VALUES(?,?,?,?,?,?,?)"
  Set INSCBCSUBREP_QR = rdb.CreateQuery("INSCBCSUBREP_QRHL", str)
  
  str = "SELECT * FROM DISLIB.DCBCSUBREP WHERE REQNO=?"
  Set SELCBCSUBREP_QR = rdb.CreateQuery("SELCBCSUBREP_QRhl", str)
  
  'STR="UPDATE DISLIB.DCBCSUBREP WHERE REQ
  '' hemato interface queries
  str = "select * from dislib.dhemcyto1 where reqno = ?"
  Set SELDHEMCYTO_QR = rdb.CreateQuery("seledhemcyto1_qr", str)
  
  '********** Shraddha Pandit 22/09/09 **********
  txt_NEW = NEWFLAG
  txt_VIEW = viewflag
  txt_SAVE = flagsave
    
  '**********************************************
  
  str = "select * from interface.hinterface where reqno = ? and iflag = '2' "
  Set selhinterface_qr = rdb.CreateQuery("selhinterface1_qr", str)
  str = "select * FROM interface.cinterface where reqno = ? and iflag = '2' "
  Set selcinterface_qr = rdb.CreateQuery("selcinterface1_rq", str)
  str = "UPDATE DISLIB.DREQUISITN SET LATERALITS=?, METASTASIS=? WHERE REQNO=?"
  Set UPDDREQDRNAME_QR = rdb.CreateQuery("UPDDREQDRNAME_QR", str)
  str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO=?"
  Set SELDREQDRNAME_QR = rdb.CreateQuery("SELDREQDRNAME_QR", str)
  str = "insert into dislib.dcbcsmrep1 values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
  Set inscbcsm_qr = rdb.CreateQuery("inscbcsdm1", str)
  str = "select * from dislib.dcbcsmrep1 where reqno = ?"
  Set selcbcsm_qr = rdb.CreateQuery("selcbcsmqr1", str)
  str = "UPDATE DISLIB.DCBCSMREP1 SET"
  str = str + " CD2 = ?, "
  str = str + " CD3 = ?, "
  str = str + " CYTOCD3 = ?, "
  str = str + " CD4 = ?, "
  str = str + " CD5 = ?, "
  str = str + " CD7 = ?, "
  str = str + " CD8 = ?, "
  str = str + " CD10 = ?, "
  str = str + " CD11C = ?, "
  str = str + " CD13 = ?, "
  str = str + " CD14 = ?, "
  str = str + " CD15 = ?, "
  str = str + " CD16 = ?, "
  str = str + " CD19 = ?, "
  str = str + " CD20 = ?, "
  str = str + " CD22 = ?, "
  str = str + " CYTOCD22 = ?, "
  str = str + " CD23 = ?, "
  str = str + " CD25 = ?, "
  str = str + " CD33 = ?, "
  str = str + " CD34 = ?, "
  str = str + " CD38 = ?, "
  str = str + " CD41 = ?, "
  str = str + " CD42 = ?, "
  str = str + " CD56 = ?, "
  str = str + " CD61 = ?, "
  str = str + " CYTOCD68 = ?, "
  str = str + " CYTOCD79A = ?, "
  str = str + " CD79B = ?, "
  str = str + " CD103 = ?, "
  str = str + " CD117 = ?, "
  str = str + " CD138 = ?, "
  str = str + " ANTIMPO = ?, "
  str = str + " CYTOMU = ?, "
  str = str + " DBA44 = ?, "
  str = str + " FMC7 = ?, "
  str = str + " GLYCOPHOIN = ?, "
  str = str + " HLADR = ?, "
  str = str + " KAPPA = ?, "
  str = str + " IGM = ?, "
  str = str + " LAMDA = ?, "
  str = str + " TDT = ?, "
  str = str + " SMIG = ?, "
  str = str + "CLND1=?, "
  str = str + "TCRAB=?, "
  str = str + "TCRGD=?, "
  str = str + "SMPCTG=?, "
  str = str + "OPTLSP=?, "
  str = str + "OPTSCP=? "
  str = str + " WHERE REQNO = ?"
  Set updcbcsm_qr = rdb.CreateQuery("UPDCBCSC1", str)
  str = "update interface.cinterface set iflag = '3' where reqno = ?"
  Set updcinterface_qr = rdb.CreateQuery("updcinter", str)
  
  str = "update interface.hinterface set iflag = '3' where reqno = ?"
  Set updhinterface_qr = rdb.CreateQuery("updhinter", str)
  
     
  
  '' hemato interface queries ends
   '''' FOR SAVING COAGULATION PROFILE
   str = "INSERT INTO DISLIB.DHEMCPDET (REQNO, PROCCODE, OBSERVALUE,OBSERTYPE, CONTROL, ENTERBY, UOM, NRANGE, MACHINE)VALUES(?,?,?,?,?,?,?,?,?)"
   Set INSCP_QR = rdb.CreateQuery("", str)
   
   'str = "SELECT * FROM DISLIB.DHEMCPRNG WHERE PROCCODE=?"
   str = "SELECT * FROM DISLIB.DHEMCPRNG WHERE PROCCODE=?"
   Set SELCPRNG_QR = rdb.CreateQuery("", str)
   str = "SELECT * FROM DISLIB.DHEMCPDET WHERE REQNO=? AND PROCCODE=?"
   
    
   Set SELCPDET_QR = rdb.CreateQuery("", str)
   str = "UPDATE DISLIB.DHEMCPDET SET OBSERVALUE=?,OBSERTYPE=?,CONTROL=?,ENTERBY=?, MACHINE=? WHERE REQNO=? AND PROCCODE=?"
   Set UPDCPDET_QR = rdb.CreateQuery("", str)
   str = "SELECT * FROM DISLIB.DHEMCPDET WHERE REQNO=?"
   Set SELCPDETVIEW_QR = rdb.CreateQuery("", str)
   
   
   str = "SELECT * FROM dISLIB.DHEMCPDET WHERE REQNO=?"
   Set SEL_DHEMCPDET_QR = rdb.CreateQuery("", str)
  End If
 
 Command1.Visible = False
 SSTab1.Tab = 0
     
   cmd_mdc.Enabled = False
 
 '***** Added by Shraddha on 27/04/2011 *****
    cmb_finalby.clear
    Dim STRSQL As String
    STRSQL = "SELECT * FROM DISLIB.DSECURITY, DISLIB.REPWRITERS Where USER_ID=USR_ID AND (( MODMO='5')OR ( USRTYPE='10')) "
    Set SEL_USER_QR = rdb.CreateQuery("", STRSQL)
    
    Set SEL_USER_RS = SEL_USER_QR.OpenResultset(1, 1)
    
    'Set SEL_USER_RS = rdb.OpenResultset(STR, rdOpenStatic, rdConcurValues)
    
    While Not SEL_USER_RS.EOF
        cmb_finalby.AddItem (Trim(SEL_USER_RS("USERNAME")))
        SEL_USER_RS.MoveNext
    Wend
 '*******************************************
 
Exit Sub
'MSG:
'MsgBox Err.Description
End Sub








Private Sub lbl_time_Click()

End Sub

Public Sub MaskEdBox1_LostFocus()
Dim CASENO As String
On Error GoTo MSG:

'If Len(Trim(MaskEdBox1.Text)) = "" Then
If Trim(MaskEdBox1.Text) = "" Then 'commented & added by Shraddha Pandit on 06/04/09
Exit Sub
End If


If Opt_patient.Value = True Then
        CASENO = MaskEdBox1.Text
       '' MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3)) ''commented by Monika on 05/12/2022
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
        If PREPATH_FLAG = False Then
            PREPATH_QR(0) = MaskEdBox1.FormattedText
            Set PREPATH_RS = PREPATH_QR.OpenResultset(1, 3)
            PREPATH_FLAG = True
        Else
            PREPATH_QR(0) = MaskEdBox1.FormattedText
            PREPATH_RS.MoreResults
            PREPATH_RS.Requery
        End If
        If Not PREPATH_RS.EOF Then
            Do While Not PREPATH_RS.EOF
                CMB_PREPATHNO.AddItem req_chglen(PREPATH_RS("REQNNO")) & " " & "$" & PREPATH_RS("PATHNO")
                PREPATH_RS.MoveNext
            Loop
        End If
         If NEWFLAG = True Then
            If REQFLAG1 = False Then
                req_qr(0) = MaskEdBox1.FormattedText
                req_qr(1) = Trim(txt_modcode.Text)
                Set req_rs = req_qr.OpenResultset(1, 3)
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
                    'cmb_reqno.ListIndex = 0 'Commented by Shraddha on 10/09/10
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
                
                btnenabl ("INI")
                Exit Sub
            End If
        
        End If
End If
    If opt_staff.Value = True Then
        CASENO = MaskEdBox1.Text
        If PREPATH_FLAG = False Then
            PREPATH_QR(0) = Trim(MaskEdBox1.FormattedText)
            Set PREPATH_RS = PREPATH_QR.OpenResultset(1, 3)
            PREPATH_FLAG = True
        Else
            PREPATH_QR(0) = Trim(MaskEdBox1.FormattedText)
            PREPATH_RS.MoreResults
            PREPATH_RS.Requery
        End If
        If Not PREPATH_RS.EOF Then
            Do While Not PREPATH_RS.EOF
                CMB_PREPATHNO.AddItem req_chglen(PREPATH_RS("REQNNO")) & " " & "$" & PREPATH_RS("PATHNO")
                PREPATH_RS.MoveNext
            Loop
        End If
        'MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))
         If NEWFLAG = True Then
            If REQFLAG1 = False Then
                req_qr(0) = Trim(MaskEdBox1.FormattedText)
                req_qr(1) = MODALITY1
                Set req_rs = req_qr.OpenResultset(1, 3)
            Else
                req_qr(0) = Trim(MaskEdBox1.FormattedText)
                req_qr(1) = MODALITY1
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
                btnenabl ("INI")
                Exit Sub
            End If
        
        End If
End If
 Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub




Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
    lbl_caseno.Caption = "Case No."
   ' MaskEdBox1.Mask = "??>/#####"
    MaskEdBox1.Enabled = True
    MaskEdBox1.SetFocus
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    lbl_caseno.Caption = "  CC No."
    MaskEdBox1.Mask = "########"
    MaskEdBox1.Enabled = True
    MaskEdBox1.SetFocus
End If
End Sub

Public Sub STAFF()
 If NEWFLAG = True Then
            NEWFLAG = False
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
                If REFDR_RS("COMMITFLAG") = "Y" Then
                    MsgBox "Report is commited , You can not Modify"
                    cmd_modify.Enabled = False
                    FLDENABL ("ADDENDUM")
                    ADDENDUMFALG = True
                    cmd_save.Enabled = True
                End If
                cmb_enterby.Text = "" & REFDR_RS("ENTERBY")
                'CMB_PROVISIONAL.Text = REFDR_RS("PREPORTBY")
                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
                    cmb_finalby.Text = ""
                Else
                    cmb_finalby.Text = "" & Trim(REFDR_RS("FREPORTBY"))
                End If
                If IsNull(Trim(REFDR_RS("PREPORTBY"))) Then
                    regcbo.Text = ""
                Else
                    regcbo.Text = "" & Trim(REFDR_RS("PREPORTBY"))
                End If
            End If

                
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = Trim(MaskEdBox1.FormattedText)
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = Trim(MaskEdBox1.FormattedText)
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If
          
          
            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = True
                txt_bedno.Visible = True
                lbl_ward.Visible = True
                'lbl_bed.Visible = True
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_status.Text = "ST"
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                'lbl_bed.Visible = False
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
                TXT_PATH.Text = Trim(selpathdrequisitn_rs("PATHNO"))
        
            End If
        
            If SELTA1012FLAG1 = False Then
                selta1012_qr(0) = Trim(MaskEdBox1.FormattedText)
                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                SELTA1012FLAG1 = True
            Else
                selta1012_qr(0) = Trim(MaskEdBox1.FormattedText)
                selta1012_rs.MoreResults
                selta1012_rs.Requery
            End If
            If selta1012_rs.RowCount > 0 Then
                'txt_procedure.Text = selta1012_rs("prvdiagdes")
            End If
                       
           
        

        
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
            
            '******
          INTERF
        
        End If
        If viewflag = True Then
        
           
            
            viewflag = False
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
                txt_category.Text = Trim(SELDREQSTAFFDRS("RELATION"))
            End If
        
            If REFDRFLAG1 = True Then
                REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
                REFDR_QR(1) = Trim(req_chglen(cmb_reqno.Text))
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = Trim(MaskEdBox1.FormattedText)
                REFDR_QR(1) = req_chglen(cmb_reqno.Text)
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If
        
'            If REFDR_RS.RowCount > 0 Then
'                txt_refdoc.Text = REFDR_RS("REFBYDOC")
'                If REFDR_RS("COMMITFLAG") = "Y" Then
'                    MsgBox "Final Report is Generated, You cann't Modify the Final Report"
'                    ADDENDUMFALG = True
'                End If
'                cmb_enterby.Text = REFDR_RS("ENTERBY")
'                'cmb_provisionalby.Text = REFDR_RS("PREPORTBY")
'                If IsNull(Trim(REFDR_RS("FREPORTBY"))) Then
'                    cmb_finalby.Text = ""
'                Else
'                    cmb_finalby.Text = REFDR_RS("FREPORTBY")
'                End If
'                If IsNull(Trim(REFDR_RS("PREPORTBY"))) Then
'                    regcbo.Text = ""  ' cmb_finalby.Text = ""
'                Else
'                    regcbo.Text = REFDR_RS("PREPORTBY")
'                End If
'            End If
            
            
            
          
            If SELMA104FLAG1 = False Then
                selma104_qr(0) = Trim(MaskEdBox1.FormattedText)
                Set selma104_rs = selma104_qr.OpenResultset(1, 3)
                SELMA104FLAG1 = True
            Else
                selma104_qr(0) = Trim(MaskEdBox1.FormattedText)
                selma104_rs.MoreResults
                selma104_rs.Requery
            End If
          
          
            If selma104_rs.RowCount > 0 Then
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                'lbl_bed.Visible = False
                txt_ward.Text = selma104_rs("WARD_NAME")
                txt_bedno.Text = selma104_rs("bed_no")
                txt_status.Text = "IP"
            Else
                txt_ward.Visible = False
                txt_bedno.Visible = False
                lbl_ward.Visible = False
                'lbl_bed.Visible = False
                txt_status.Text = "ST"
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
                TXT_PATH.Text = selpathdrequisitn_rs("PATHNO")
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
            ''''''SUNDEEP FOR VIEW OF hl**********************
            If SELDBMAREPORTFLAG = False Then
                SELDBMAREPORT_QR(0) = req_chglen(cmb_reqno.Text)
                Set SELDBMAREPORT_RS = SELDBMAREPORT_QR.OpenResultset(1, 3)
                SELDBMAREPORTFLAG = True
            Else
                SELDBMAREPORT_QR(0) = req_chglen(cmb_reqno.Text)
                SELDBMAREPORT_RS.MoreResults
                SELDBMAREPORT_RS.Requery
            End If
            ENABL ("DIS")
            VIEW_CBC
            If SELDBMAREPORT_RS.RowCount > 0 Then
                    '**** cbc counter
                    If SELCBCSUBREPFLAG = False Then
                        SELCBCSUBREP_QR(0) = req_chglen(cmb_reqno.Text)
                        Set SELCBCSUBREP_RS = SELCBCSUBREP_QR.OpenResultset(1, 3)
                        SELCBCSUBREPFLAG = False
                    Else
                        SELCBCSUBREP_QR(0) = req_chglen(cmb_reqno.Text)
                        SELCBCSUBREP_RS.MoreResults
                        SELCBCSUBREP_RS.Requery
                    End If
                    If SELCBCSUBREP_RS.RowCount > 0 Then
                        SELCBCSUBREP_RS.MoveLast
                        TXT_HB.Text = " " & SELCBCSUBREP_RS("hb")
                        TXT_RBC.Text = "" & SELCBCSUBREP_RS("rbc")
                        TXT_TWBC.Text = " " & SELCBCSUBREP_RS("tlc")
                        txt_platelets.Text = " " & SELCBCSUBREP_RS("platelets")
                    End If
                    
           ''*******ADDENDUM*********************************
            'If SELDADDENDUMFLAG = False Then
             '   SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
              '  Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
               ' SELDADDENDUMFLAG = True
            'Else
            '' HERE
                SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
                SELDADDENDUM_RS.MoreResults
                SELDADDENDUM_RS.Requery
            'End If
            If SELDADDENDUM_RS.RowCount > 0 Then
                ADDENDUM = ""
                While Not SELDADDENDUM_RS.EOF
                    ADDENDUM = ADDENDUM & " ADDENDUM NO " & SELDADDENDUM_RS("ADDEND_NO") & Chr(13) & SELDADDENDUM_RS("ADD_MAX_LN") & Chr(13) & SELDADDENDUM_RS("ADDEND_DT") & " " & SELDADDENDUM_RS("ADDEND_TM") & " " & SELDADDENDUM_RS("USER_ID")
                    'cmb_addendumby.AddItem SELDADDENDUM_RS("USER_ID")
                    cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
                SELDADDENDUM_RS.MoveNext
                Wend
                txt_addendum.Text = Trim(ADDENDUM)
                ADDENDUMFALG = True
            End If
            '' ADD
            '''  surface marker data fetching for view
            selcbcsm_qr(0) = req_chglen(cmb_reqno.Text)
               Set selcbcsm_rs = selcbcsm_qr.OpenResultset(1, 1)
                If selcbcsm_rs.RowCount > 0 Then
                    SURFACEFORM.TXT_ANTIMPO = selcbcsm_rs("antimpo") '
                    SURFACEFORM.TXT_CD10 = selcbcsm_rs("cd10") '" "
                    SURFACEFORM.txt_cd103 = selcbcsm_rs("cd103") '  " "
                    SURFACEFORM.TXT_CD117 = selcbcsm_rs("cd117") ' " "
                    SURFACEFORM.TXT_CD11C = selcbcsm_rs("cd11c")  '" "
                    
                    SURFACEFORM.txt_cd13 = selcbcsm_rs("cd13")  '" "
                    SURFACEFORM.TXT_CD138 = selcbcsm_rs("cd138")  '" "
                    SURFACEFORM.txt_cd14 = selcbcsm_rs("cd14")  '" "
                    SURFACEFORM.txt_CD15 = selcbcsm_rs("cd15")  '" "
                    SURFACEFORM.TXT_CD19 = selcbcsm_rs("cd19")  ' " "
                    SURFACEFORM.TXT_CD2 = selcbcsm_rs("cd2")   '" "
                    SURFACEFORM.TXT_CD20 = selcbcsm_rs("cd20")  ' ""
                    SURFACEFORM.TXT_CD22 = selcbcsm_rs("cd22")  ' ""
                    SURFACEFORM.TXT_CD23 = selcbcsm_rs("cd23")  ' ""
                    SURFACEFORM.TXT_CD25 = selcbcsm_rs("cd25")  '""
                    SURFACEFORM.TXT_CD3 = selcbcsm_rs("cd3")  ' " "
                    SURFACEFORM.txt_CD33 = selcbcsm_rs("cd33")  '""
                    SURFACEFORM.TXT_CD34 = selcbcsm_rs("cd34")  '" "
                    SURFACEFORM.txt_cd36 = selcbcsm_rs("cd16")  '" "
                    SURFACEFORM.TXT_CD38 = selcbcsm_rs("cd38") ' " "
                    SURFACEFORM.TXT_CD4 = selcbcsm_rs("cd4")  ' " "
                    SURFACEFORM.txt_cd41 = selcbcsm_rs("cd41")  ' " "
                    SURFACEFORM.TXT_CD42 = selcbcsm_rs("cd42")  ' " "
                    SURFACEFORM.TXT_CD5 = selcbcsm_rs("cd5")  ' " "
                    SURFACEFORM.TXT_CD56 = selcbcsm_rs("cd56")  '" "
                    SURFACEFORM.TXT_CD61 = selcbcsm_rs("cd61")  '" "
                    SURFACEFORM.TXT_CD68 = selcbcsm_rs("cytocd68")  '" "
                    SURFACEFORM.TXT_CD7 = selcbcsm_rs("cd7")   '" "
                    SURFACEFORM.txt_CD79B = selcbcsm_rs("cd79b")  '" "
                    SURFACEFORM.TXT_CD8 = selcbcsm_rs("cd8")  '""
                    SURFACEFORM.txt_cytocd221(1) = selcbcsm_rs("cytocd22")  '" "
                    SURFACEFORM.TXT_CYTOCD3 = selcbcsm_rs("cytocd3")  ' " "
                    SURFACEFORM.txt_cytocd79a = selcbcsm_rs("cytocd79a")  '" "
                    SURFACEFORM.TXT_CYTOMU = selcbcsm_rs("cytomu")  ' " "
                    SURFACEFORM.TXT_DBA44 = selcbcsm_rs("dba44")  '" "
                    SURFACEFORM.TXT_FMC7 = selcbcsm_rs("fmc7")  '" "
                    SURFACEFORM.txt_Glycophorina = selcbcsm_rs("glycophoin")  '" "
                    SURFACEFORM.TXT_HLADR = selcbcsm_rs("hladr")  '" "
                    SURFACEFORM.txt_igm = selcbcsm_rs("igm")  '" "
                    SURFACEFORM.txt_kappa = selcbcsm_rs("kappa")  '" "
                    SURFACEFORM.txt_lanbda = selcbcsm_rs("lamda")  '" "
                    SURFACEFORM.txt_smig = selcbcsm_rs("smig")  '" "
                    SURFACEFORM.TXT_TDT = selcbcsm_rs("tdt")  ' " "
                              
                        If selcbcsm_rs("optlsp") = "Y" Then
                            SURFACEFORM.opt_lsp.Value = True
                        ElseIf selcbcsm_rs("optlsp") = "N" Then
                            SURFACEFORM.opt_lsp.Value = False
                        End If
                    
                        If selcbcsm_rs("optscp") = "Y" Then
                            SURFACEFORM.opt_ssp.Value = True
                        ElseIf selcbcsm_rs("optscp") = "N" Then
                            SURFACEFORM.opt_ssp.Value = False
                        End If
                        
                        
                 
                          
                        SELDREQDRNAME_QR(0) = req_chglen(cmb_reqno.Text)
                        Set SELDREQDRNAME_RS = SELDREQDRNAME_QR.OpenResultset(1, 3)
                        
                        If SELDREQDRNAME_RS.RowCount > 0 Then
                               If Trim(SELDREQDRNAME_RS("LATERALITS")) = "Dr. Ashok Kumar" Then
                                   SURFACEFORM.opt_drak.Value = True
                                   SURFACEFORM.opt_mryb.Value = False
                                   SURFACEFORM.opt_sitaram.Value = False
                               ElseIf Trim(SELDREQDRNAME_RS("METASTASIS")) = "Mr. Y. Badrinath" Then
                                   SURFACEFORM.opt_drak = False
                                   SURFACEFORM.opt_mryb.Value = True
                                   SURFACEFORM.opt_sitaram.Value = False
                                   '***** Added by Shradha on 05/05/10 *****
                                ElseIf Trim(SELDREQDRNAME_RS("METASTASIS")) = "Mr. Sitaram" Then
                                    SURFACEFORM.opt_drak.Value = False
                                    SURFACEFORM.opt_mryb.Value = False
                                    SURFACEFORM.opt_sitaram.Value = True
                                End If
                        End If
                
                
                
                 End If
                          
           '''  surface marker data fetching for view ends
            '' ADD
            ''*******ADDENDUM*********************************

                    
                    
                    'txt_neutrophils.Text = " " & SELDBMAREPORT_RS("neutrophil")
                    'txt_lymphocytes.Text = " " & SELDBMAREPORT_RS("lymphocyte")
                    'txt_monocytes.Text = " " & SELDBMAREPORT_RS("monocytes")
                    'txt_eosinophils.Text = " " & SELDBMAREPORT_RS("eosinophil")
                    'txt_basophils.Text = " " & SELDBMAREPORT_RS("basophils")
                    'txt_immaturegranulocytes.Text = " " & SELDBMAREPORT_RS("immgranulo")
                    'TXT_HCT.Text = " " & SELDBMAREPORT_RS("hct")
                    'TXT_MCH.Text = " " & SELDBMAREPORT_RS("mch")
                    'TXT_MCV.Text = " " & SELDBMAREPORT_RS("mcv")
                    
                    'TXT_MCHC.Text = " " & SELDBMAREPORT_RS("mchc")
                    'TXT_MPV.Text = " " & SELDBMAREPORT_RS("mpv")
                    'TXT_RDW.Text = " " & SELDBMAREPORT_RS("rdw")
                    
                    
                    '''DONE BY KANCHEN
                    VIEW_CBC
                    TXT_ABNORMAL.Text = " " & SELDBMAREPORT_RS("abnormal")
                    '*** cbc counter
                    
                    
                    
                    '*** cbc manual
                    
                    txt_mneutrophils.Text = " " & SELDBMAREPORT_RS("neutrophi1")
                    txt_mlymphocytes.Text = " " & SELDBMAREPORT_RS("lymphocyt1")
                    txt_prolympho.Text = " " & SELDBMAREPORT_RS("pLTmorpho") '13/11/08 Shraddha Pandit
                    txt_mmonocytes.Text = " " & SELDBMAREPORT_RS("monocyte1")
                    txt_meosinophils.Text = " " & SELDBMAREPORT_RS("eosinophi1")
                    txt_mbasophills.Text = " " & SELDBMAREPORT_RS("basophill1")
                    txt_blasts.Text = " " & SELDBMAREPORT_RS("blasts")
                    TXT_PROMYELOCYTESWBC.Text = " " & SELDBMAREPORT_RS("promyelocy")
                    txt_myelocytes.Text = " " & SELDBMAREPORT_RS("myelocytes")
                    txt_metamyelocytes.Text = " " & SELDBMAREPORT_RS("metamyeloc")
                    txt_plasmacells.Text = " " & SELDBMAREPORT_RS("plasmacell")
                    txt_alymphocytes.Text = " " & SELDBMAREPORT_RS("atlymphee")
                    txt_prolymphocytes.Text = " " & SELDBMAREPORT_RS("prolymphoc")
                    txt_Sezarycells.Text = " " & SELDBMAREPORT_RS("sezarycell")
                    txt_hairycells.Text = " " & SELDBMAREPORT_RS("hairycells")
                    txt_othercells.Text = " " & SELDBMAREPORT_RS("othercells")
                    TXT_NORMOBLASTS.Text = " " & SELDBMAREPORT_RS("normboblas")
                    TXT_OTHERCELLSRBC.Text = " " & SELDBMAREPORT_RS("othercellr")
                    txt_rbcmorphology.Text = " " & SELDBMAREPORT_RS("rbcmorpho")
                    TXT_LP.Text = " " & SELDBMAREPORT_RS("FLD1") 'Shraddha Pandit 17/10/08
                    cmbmpo.Text = " " & SELDBMAREPORT_RS("mpoval")
                    txt_reticcount.Text = " " & SELDBMAREPORT_RS("reticcount")
                    txt_malarialparasite.Text = " " & SELDBMAREPORT_RS("mparasite")
                    '************************************
'                    txt_malarialparasite.Text = " " & DCBCREPA_RS("MALERIA") 'Shraddha Pandit 20/10/08
'                    txt_othertype.Text = " " & DCBCREPA_RS("FLD2") 'Shraddha Pandit 21/10/08
'                    '************************************
                    txt_esr.Text = " " & SELDBMAREPORT_RS("esr") '18/11/08
                    'TXT_MORPHOLOGY.Text = " " & SELDBMAREPORT_RS("pmorpholog")
                    CMB_MORPHO.Text = " " & SELDBMAREPORT_RS("pmorpholog") 'Shraddha 13/11/08
                    TXT_ADEQUACY.Text = " " & SELDBMAREPORT_RS("fld1")
                    
                    cmbmpo.Text = " " & SELDBMAREPORT_RS("MPO")
                    VIEWDCBCREPA
                    '*** cbc manual
                    
                    '*** coagulation profile
                    txt_inr.Text = " " & SELDBMAREPORT_RS("bt")
                    TXT_PT(0).Text = " " & SELDBMAREPORT_RS("pt")
                    TXT_PTTK.Text = " " & SELDBMAREPORT_RS("pttk")
                    txt_ptctl.Text = " " & SELDBMAREPORT_RS("platelet1")
                    TXT_FIBRINOGEN.Text = " " & SELDBMAREPORT_RS("fiBrinogen")
                    TXT_FDPDMIER.Text = " " & SELDBMAREPORT_RS("fdpddimer")
                    TXT_OTHERS.Text = " " & SELDBMAREPORT_RS("others")
                    TXT_TT.Text = " " & SELDBMAREPORT_RS("tt")
                    '*** cogulation profile
                    
                    '*** cytochemistry
                    txt_lapscore.Text = " " & SELDBMAREPORT_RS("lap")
                    TXT_MPO.Text = " " & SELDBMAREPORT_RS("mpo")
                    cmb_nse.Text = " " & SELDBMAREPORT_RS("sbb")
                    TXT_PAS.Text = " " & SELDBMAREPORT_RS("pSA")
                    txt_promonocyte.Text = " " & SELDBMAREPORT_RS("pSA")
                    TXT_NSE.Text = " " & SELDBMAREPORT_RS("nse")
                    TXT_TRAP.Text = " " & SELDBMAREPORT_RS("traD")
                    
                    'cytochemistry
                    '''***FOR COAGULATION PROFILE OF STAFF
                    SELCPDETVIEW_QR(0) = req_chglen(cmb_reqno.Text)
                 Set SELCPDETVIEW_RS = SELCPDETVIEW_QR.OpenResultset(2, 1)
                 If Not SELCPDETVIEW_RS.EOF Then
                    Do While Not SELCPDETVIEW_RS.EOF
                        If Trim(SELCPDETVIEW_RS("PROCCODE")) = "PT" Then
                            TXT_PT(0).Text = Trim(SELCPDETVIEW_RS("OBSERVALUE"))
                            txt_ptctl.Text = Trim(SELCPDETVIEW_RS("CONTROL"))
                        
                        End If
                        If Trim(SELCPDETVIEW_RS("PROCCODE")) = "INR" Then
                            txt_inr.Text = Trim(SELCPDETVIEW_RS("OBSERVALUE"))
                            
                        
                        End If
                        If Trim(SELCPDETVIEW_RS("PROCCODE")) = "APTT" Then
                            TXT_PTTK.Text = Trim(SELCPDETVIEW_RS("OBSERVALUE"))
                            txt_apttctl.Text = Trim(SELCPDETVIEW_RS("CONTROL"))
                        
                        End If
                        If Trim(SELCPDETVIEW_RS("PROCCODE")) = "TT" Then
                            TXT_TT.Text = Trim(SELCPDETVIEW_RS("OBSERVALUE"))
                            txt_ttctl.Text = Trim(SELCPDETVIEW_RS("CONTROL"))
                        
                        End If
                        If Trim(SELCPDETVIEW_RS("PROCCODE")) = "FIBRINOGEN" Then
                            TXT_FIBRINOGEN.Text = Trim(SELCPDETVIEW_RS("OBSERVALUE"))
                        
                        End If
                        If Trim(SELCPDETVIEW_RS("PROCCODE")) = "D-DIMER" Then
                            TXT_FDPDMIER.Text = Trim(SELCPDETVIEW_RS("OBSERVALUE"))
                        
                        End If
'                        If Trim(SELCPDETVIEW_RS("MACHINE")) = "ACL 100" Then
'                            Option1.Value = True
'                            Option2.Value = False
'                            Option3.Value = False
'                        ElseIf Trim(SELCPDETVIEW_RS("MACHINE")) = "ACL 200" Then
'                            Option1.Value = False
'                            Option2.Value = True
'                            Option3.Value = False
'                             ElseIf Trim(SELCPDETVIEW_RS("MACHINE")) = "ELP 001" Then
'                            Option1.Value = False
'                            Option2.Value = False
'                            Option3.Value = True
'                            Else
'                            cmb_machname.Text = Trim(SELCPDETVIEW_RS("MACHINE"))
'                        End If

                            If IsNull(SELCPDETVIEW_RS("MACHINE")) = False Then
                            cmb_machname.Text = Trim(SELCPDETVIEW_RS("MACHINE"))
                            End If

                        SELCPDETVIEW_RS.MoveNext
                    Loop
                 End If
                    
                    
                    
                    txt_impression.Text = " " & SELDBMAREPORT_RS("IMPRESSION")
            End If
 
            'End If
            
            '*************************************************
            
        End If
End Sub

Public Sub VIEWDCBCREPA()
           
    If SELDCBCFLAG = False Then
        SELDCBCREPA_QR(0) = req_chglen(cmb_reqno.Text)
        Set SELDCBCREPA_RS = SELDCBCREPA_QR.OpenResultset(1, 3)
        SELDCBCFLAG = True
    Else
        SELDCBCREPA_QR(0) = req_chglen(cmb_reqno.Text)
        SELDCBCREPA_RS.MoreResults
        SELDCBCREPA_RS.Requery
    End If
    
    If SELDCBCREPA_RS.RowCount > 0 Then
        txt_crc.Text = SELDCBCREPA_RS("CRC")
        txt_rpi.Text = SELDCBCREPA_RS("RPI")
        txt_empc.Text = SELDCBCREPA_RS("EMPCNT")
        txt_othertype = SELDCBCREPA_RS("fld2")
        'txt_promonocyte.Text = SELDCBCREPA_RS("fld2") '19/11/08
        txt_malarialparasite.Text = SELDCBCREPA_RS("maleria")
        If IsNull(SELDCBCREPA_RS("fld3")) = False Then
            txt_cbc_comment.Text = SELDCBCREPA_RS("fld3") 'Shraddha 01/03/11
        End If
        
        If IsNull(SELDCBCREPA_RS("fld7")) = False Then
            TXT_DNAINDEX.Text = SELDCBCREPA_RS("fld7") 'SHWETA 10-02-2014
        End If
        
        If IsNull(SELDCBCREPA_RS("fld6")) = False Then
            TXT_DNACLOIDY.Text = SELDCBCREPA_RS("fld6") 'SHWETA 10-02-2014
        End If
        
    End If
End Sub

Public Sub DCBCREPA()

'''' Done by Kanchan
  
  If flagsave = 1 Then

        INSDCBCREPA_QR(0) = req_chglen(cmb_reqno.Text)
    
        If Len(Trim(txt_crc.Text)) > 0 Then
            INSDCBCREPA_QR(1) = Trim(txt_crc.Text)
        Else
            INSDCBCREPA_QR(1) = "  "
        End If
        
        If Len(Trim(txt_rpi.Text)) > 0 Then
            INSDCBCREPA_QR(2) = Trim(txt_rpi.Text)
        Else
            INSDCBCREPA_QR(2) = "  "
        End If
        
        If Len(Trim(txt_empc.Text)) > 0 Then
            INSDCBCREPA_QR(3) = Trim(txt_empc.Text)
        Else
            INSDCBCREPA_QR(3) = "  "
        End If
        
        If Len(Trim(txt_malarialparasite.Text)) > 0 Then
            INSDCBCREPA_QR(4) = Trim(txt_malarialparasite.Text)
        Else
            INSDCBCREPA_QR(4) = " "
        End If
        
        If Len(Trim(txt_othertype.Text)) > 0 Then
            INSDCBCREPA_QR(5) = Trim(txt_othertype.Text)
        Else
            INSDCBCREPA_QR(5) = 0
        End If
        
        '***** Shraddha 01/03/11 *****
        If Len(Trim(txt_cbc_comment.Text)) > 0 Then
            INSDCBCREPA_QR(6) = Trim(txt_cbc_comment.Text)
        Else
            INSDCBCREPA_QR(6) = ""
        End If
        '*****************************
    
    
    '***** SHWETA 10-02-2014 *****
        If Len(Trim(TXT_DNAINDEX.Text)) > 0 Then
            INSDCBCREPA_QR(7) = Trim(TXT_DNAINDEX.Text)
        Else
            INSDCBCREPA_QR(7) = ""
        End If
        
            If Len(Trim(TXT_DNACLOIDY.Text)) > 0 Then
            INSDCBCREPA_QR(8) = Trim(TXT_DNACLOIDY.Text)
        Else
            INSDCBCREPA_QR(8) = ""
        End If
        
        '*****************************
    
    INSDCBCREPA_QR.Execute

  ElseIf flagsave = 2 Then

    UPDDCBCREPA_QR(0) = Trim(txt_crc.Text)
    UPDDCBCREPA_QR(1) = Trim(txt_rpi.Text)
    UPDDCBCREPA_QR(2) = Trim(txt_empc.Text)
    UPDDCBCREPA_QR(3) = Trim(txt_malarialparasite.Text)
    UPDDCBCREPA_QR(4) = Trim(txt_othertype.Text)
    UPDDCBCREPA_QR(5) = Trim(txt_cbc_comment.Text) 'Added on 01/03/11 by Shraddha
    UPDDCBCREPA_QR(6) = Trim(TXT_DNAINDEX.Text) 'Added on 10-02-2014 by SHWETA
    UPDDCBCREPA_QR(7) = Trim(TXT_DNACLOIDY.Text) 'Added on 10-02-2014 by SHWETA
    UPDDCBCREPA_QR(8) = req_chglen(cmb_reqno.Text)
    UPDDCBCREPA_QR.Execute

  End If

End Sub
Public Sub assign_new()
    INSDBMAREPORT_QR(0) = req_chglen(cmb_reqno.Text)
    If Len(Trim(TXT_HB.Text)) > 0 Then
        INSDBMAREPORT_QR(1) = Trim(TXT_HB.Text)
    Else
        INSDBMAREPORT_QR(1) = "  "
    End If
    If Len(Trim(TXT_RBC.Text)) > 0 Then
        INSDBMAREPORT_QR(2) = Trim(TXT_RBC.Text)
    Else
        INSDBMAREPORT_QR(2) = "  "
    End If
    If Len(Trim(TXT_TWBC.Text)) > 0 Then
        INSDBMAREPORT_QR(3) = Trim(TXT_TWBC.Text)
    Else
        INSDBMAREPORT_QR(3) = "  "
    End If
    If Len(Trim(txt_platelets.Text)) > 0 Then
        INSDBMAREPORT_QR(4) = Trim(txt_platelets.Text)
    Else
        INSDBMAREPORT_QR(4) = "  "
    End If
    If Len(Trim(txt_neutrophils.Text)) > 0 Then
        INSDBMAREPORT_QR(5) = Trim(txt_neutrophils.Text)
    Else
        INSDBMAREPORT_QR(5) = "  "
    End If
    If Len(Trim(txt_lymphocytes.Text)) > 0 Then
        INSDBMAREPORT_QR(6) = Trim(txt_lymphocytes.Text)
    Else
        INSDBMAREPORT_QR(6) = Trim(txt_lymphocytes.Text)
    End If
    If Len(Trim(txt_monocytes.Text)) > 0 Then
        INSDBMAREPORT_QR(7) = Trim(txt_monocytes.Text)
    Else
        INSDBMAREPORT_QR(7) = "  "
    End If
    If Len(Trim(txt_eosinophils.Text)) > 0 Then
        INSDBMAREPORT_QR(8) = Trim(txt_eosinophils.Text)
    Else
        INSDBMAREPORT_QR(8) = "  "
    End If
    If Len(Trim(txt_basophils.Text)) > 0 Then
        INSDBMAREPORT_QR(9) = Trim(txt_basophils.Text)
    Else
        INSDBMAREPORT_QR(9) = "  "
    End If
    If Len(Trim(txt_immaturegranulocytes.Text)) > 0 Then
        INSDBMAREPORT_QR(10) = Trim(txt_immaturegranulocytes.Text)
    Else
        INSDBMAREPORT_QR(10) = "  "
    End If
    If Len(Trim(TXT_HCT.Text)) > 0 Then
        INSDBMAREPORT_QR(11) = Trim(TXT_HCT.Text)
    Else
        INSDBMAREPORT_QR(11) = "  "
    End If
    If Len(Trim(TXT_MCH.Text)) > 0 Then
        INSDBMAREPORT_QR(12) = Trim(TXT_MCH.Text)
    Else
        INSDBMAREPORT_QR(12) = "  "
    End If
    If Len(Trim(TXT_MCV.Text)) > 0 Then
        INSDBMAREPORT_QR(13) = Trim(TXT_MCV.Text)
    Else
        INSDBMAREPORT_QR(13) = "  "
    End If
    If Len(Trim(TXT_MCHC.Text)) > 0 Then
        INSDBMAREPORT_QR(14) = Trim(TXT_MCHC.Text)
    Else
        INSDBMAREPORT_QR(14) = "  "
    End If
    If Len(Trim(TXT_MPV.Text)) > 0 Then
        INSDBMAREPORT_QR(15) = Trim(TXT_MPV.Text)
    Else
        INSDBMAREPORT_QR(15) = "  "
    End If
    If Len(Trim(TXT_RDW.Text)) > 0 Then
        INSDBMAREPORT_QR(16) = Trim(TXT_RDW.Text)
    Else
        INSDBMAREPORT_QR(16) = "  "
    End If
    If Len(Trim(TXT_ABNORMAL.Text)) > 0 Then
        INSDBMAREPORT_QR(17) = Trim(TXT_ABNORMAL.Text)
    Else
        INSDBMAREPORT_QR(17) = "  "
    End If
    '*** CBC COUNTER
    '*** CBC MANUAL
    If Len(Trim(txt_mneutrophils.Text)) > 0 Then
        INSDBMAREPORT_QR(18) = Trim(txt_mneutrophils.Text)
    Else
        INSDBMAREPORT_QR(18) = "  "
    End If
    If Len(Trim(txt_mlymphocytes.Text)) > 0 Then
        INSDBMAREPORT_QR(19) = Trim(txt_mlymphocytes.Text)
    Else
        INSDBMAREPORT_QR(19) = "  "
    End If
    If Len(Trim(txt_mmonocytes.Text)) > 0 Then
        INSDBMAREPORT_QR(20) = Trim(txt_mmonocytes.Text)
    Else
        INSDBMAREPORT_QR(20) = "  "
    End If
    If Len(Trim(txt_meosinophils.Text)) > 0 Then
        INSDBMAREPORT_QR(21) = Trim(txt_meosinophils.Text)
    Else
        INSDBMAREPORT_QR(21) = "  "
    End If
    If Len(Trim(txt_mbasophills.Text)) > 0 Then
        INSDBMAREPORT_QR(22) = Trim(txt_mbasophills.Text)
    Else
        INSDBMAREPORT_QR(22) = "  "
    End If
    If Len(Trim(txt_blasts.Text)) > 0 Then
        INSDBMAREPORT_QR(23) = Trim(txt_blasts.Text)
    Else
        INSDBMAREPORT_QR(23) = "  "
    End If
    If Len(Trim(TXT_PROMYELOCYTESWBC.Text)) > 0 Then
        INSDBMAREPORT_QR(24) = Trim(TXT_PROMYELOCYTESWBC.Text)
    Else
        INSDBMAREPORT_QR(24) = "  "
    End If
    If Len(Trim(txt_myelocytes.Text)) > 0 Then
        INSDBMAREPORT_QR(25) = Trim(txt_myelocytes.Text)
    Else
        INSDBMAREPORT_QR(25) = "  "
    End If
    If Len(Trim(txt_metamyelocytes.Text)) > 0 Then
        INSDBMAREPORT_QR(26) = Trim(txt_metamyelocytes.Text)
    Else
        INSDBMAREPORT_QR(26) = "  "
    End If
    If Len(Trim(txt_plasmacells.Text)) > 0 Then
        INSDBMAREPORT_QR(27) = Trim(txt_plasmacells.Text)
    Else
        INSDBMAREPORT_QR(27) = "  "
    End If
    If Len(Trim(txt_alymphocytes.Text)) > 0 Then
        INSDBMAREPORT_QR(28) = Trim(txt_alymphocytes.Text)
    Else
        INSDBMAREPORT_QR(28) = "  "
    End If
    If Len(Trim(txt_prolympho.Text)) > 0 Then
        INSDBMAREPORT_QR(29) = Trim(txt_prolympho.Text)
    Else
        INSDBMAREPORT_QR(29) = "  "
    End If
    If Len(Trim(txt_Sezarycells.Text)) > 0 Then
        INSDBMAREPORT_QR(30) = Trim(txt_Sezarycells.Text)
    Else
        INSDBMAREPORT_QR(30) = "  "
    End If
    If Len(Trim(txt_hairycells.Text)) > 0 Then
        INSDBMAREPORT_QR(31) = Trim(txt_hairycells.Text)
    Else
        INSDBMAREPORT_QR(31) = "  "
    End If
    If Len(Trim(txt_othercells.Text)) > 0 Then
        INSDBMAREPORT_QR(32) = Trim(txt_othercells.Text)
    Else
        INSDBMAREPORT_QR(32) = "  "
    End If
    If Len(Trim(TXT_NORMOBLASTS.Text)) > 0 Then
        INSDBMAREPORT_QR(33) = Trim(TXT_NORMOBLASTS.Text)
    Else
        INSDBMAREPORT_QR(33) = "  "
    End If
    If Len(Trim(TXT_OTHERCELLSRBC.Text)) > 0 Then
        INSDBMAREPORT_QR(34) = Trim(TXT_OTHERCELLSRBC.Text)
    Else
        INSDBMAREPORT_QR(34) = "  "
    End If
    If Len(Trim(txt_rbcmorphology.Text)) > 0 Then
        INSDBMAREPORT_QR(35) = Trim(txt_rbcmorphology.Text)
    Else
        INSDBMAREPORT_QR(35) = "  "
    End If
    '********** Shraddha Pandit 17/10/08 *************
    '        If Len(Trim(TXT_LP.Text)) > 0 Then
    '            INSDBMAREPORT_QR(35) = Trim(txt_LP.Text)
    '        Else
    '             INSDBMAREPORT_QR(35) = "  "
    '        End If
    '*************************************************
    If Len(Trim(txt_reticcount.Text)) > 0 Then
        INSDBMAREPORT_QR(36) = Trim(txt_reticcount.Text)
    Else
        INSDBMAREPORT_QR(36) = "  "
    End If
    If Len(Trim(txt_malarialparasite.Text)) > 0 Then
        INSDBMAREPORT_QR(37) = " " 'Trim(txt_malarialparasite.Text)
    Else
        INSDBMAREPORT_QR(37) = "  "
    End If
    If Len(Trim(txt_esr.Text)) > 0 Then '18/11/08
        INSDBMAREPORT_QR(38) = Trim(txt_esr.Text)
    Else
        INSDBMAREPORT_QR(38) = "  "
    End If
    'INSDBMAREPORT_QR(38) = "  "
    'If Len(Trim(TXT_MORPHOLOGY.Text)) > 0 Then
    If Len(Trim(CMB_MORPHO.Text)) > 0 Then
        'INSDBMAREPORT_QR(39) = Trim(TXT_MORPHOLOGY.Text)
        INSDBMAREPORT_QR(39) = Trim(CMB_MORPHO.Text)
    Else
        INSDBMAREPORT_QR(39) = "  "
    End If
            
    If Len(Trim(cmbmpo.Text)) > 0 Then
        INSDBMAREPORT_QR(49) = Trim(cmbmpo.Text)
    Else
        INSDBMAREPORT_QR(49) = "  "
    End If
        
    DCBCREPA
    
    '*** CBC MANUAL
    '*** COAGULATION PROFILE
    If Len(Trim(txt_inr.Text)) > 0 Then
        INSDBMAREPORT_QR(40) = Trim(txt_inr.Text)
    Else
        INSDBMAREPORT_QR(40) = "  "
    End If
    If Len(Trim(TXT_PT(0).Text)) > 0 Then
        INSDBMAREPORT_QR(41) = Trim(TXT_PT(0).Text)
    Else
        INSDBMAREPORT_QR(41) = "  "
    End If
    If Len(Trim(txt_ptctl.Text)) > 0 Then
        INSDBMAREPORT_QR(42) = Trim(txt_ptctl.Text) 'platelet1
    Else
        INSDBMAREPORT_QR(42) = "  "
    End If
    If Len(Trim(TXT_PTTK.Text)) > 0 Then
        INSDBMAREPORT_QR(43) = Trim(TXT_PTTK.Text)
    Else
        INSDBMAREPORT_QR(43) = "  "
    End If
    If Len(Trim(TXT_TT.Text)) > 0 Then
        INSDBMAREPORT_QR(44) = Trim(TXT_TT.Text)
    Else
        INSDBMAREPORT_QR(44) = "  "
    End If
    If Len(Trim(TXT_FIBRINOGEN.Text)) > 0 Then
        INSDBMAREPORT_QR(45) = Trim(TXT_FIBRINOGEN.Text)
    Else
        INSDBMAREPORT_QR(45) = "  "
    End If
    If Len(Trim(TXT_FDPDMIER.Text)) > 0 Then
        INSDBMAREPORT_QR(46) = Trim(TXT_FDPDMIER.Text)
    Else
        INSDBMAREPORT_QR(46) = "  "
    End If
    If Len(Trim(TXT_OTHERS.Text)) > 0 Then
        INSDBMAREPORT_QR(47) = Trim(TXT_OTHERS.Text)
    Else
        INSDBMAREPORT_QR(47) = "  "
    End If
    '*** COAGULATION PROFILE
    '*** CYTOCHEMISTRY
    If Len(Trim(txt_lapscore.Text)) > 0 Then
        INSDBMAREPORT_QR(48) = Trim(txt_lapscore.Text)
    Else
        INSDBMAREPORT_QR(48) = "  "
    End If
    'If Len(Trim(TXT_MPO.Text)) > 0 Then
    '    INSDBMAREPORT_QR(49) = Trim(TXT_MPO.Text)
    'Else
    '    INSDBMAREPORT_QR(49) = "  "
    'End If
    If Len(Trim(cmb_nse.Text)) > 0 Then
        INSDBMAREPORT_QR(50) = Trim(cmb_nse.Text)
    Else
        INSDBMAREPORT_QR(50) = "  "
    End If
    
'    If Len(Trim(TXT_PAS.Text)) > 0 Then
'        INSDBMAREPORT_QR(51) = Trim(TXT_PAS.Text)
'    Else
'        INSDBMAREPORT_QR(51) = "  "
'    End If

    If Len(Trim(txt_promonocyte.Text)) > 0 Then
        INSDBMAREPORT_QR(51) = Trim(txt_promonocyte.Text)
    Else
        INSDBMAREPORT_QR(51) = "  "
    End If
    
    If Len(Trim(TXT_NSE.Text)) > 0 Then
        INSDBMAREPORT_QR(52) = Trim(TXT_NSE.Text)
    Else
        INSDBMAREPORT_QR(52) = "  "
    End If
    If Len(Trim(TXT_TRAP.Text)) > 0 Then
        INSDBMAREPORT_QR(53) = Trim(TXT_TRAP.Text)
    Else
        INSDBMAREPORT_QR(53) = "  "
    End If
    If Len(Trim(TXT_LP.Text)) > 0 Then 'TXT_ADEQUACY.Text 20/11/08
        INSDBMAREPORT_QR(54) = Trim(TXT_LP.Text)
    Else
        INSDBMAREPORT_QR(54) = "  "
    End If
    If Len(Trim(txt_impression.Text)) > 0 Then
        'INSDBMAREPORT_QR(55) = "  "
        INSDBMAREPORT_QR(55) = Trim(txt_impression.Text)
    Else
        INSDBMAREPORT_QR(55) = "  "
    End If
        
    If Len(Trim(Text1.Text)) > 0 Then
        INSDBMAREPORT_QR(56) = Text1.Text  ' platelet morphology
    Else
        INSDBMAREPORT_QR(56) = " "
    End If
        
    '******** 16/10/08 Shraddha Pandit **********
    If Len(Trim(txt_prolympho.Text)) > 0 Then
        INSDBMAREPORT_QR(56) = Trim(txt_prolympho.Text)
    Else
        INSDBMAREPORT_QR(56) = "  "
    End If
    '****************************

    If Len(Trim(txt_apttctl.Text)) > 0 Then
        INSDBMAREPORT_QR(57) = Trim(txt_apttctl.Text)
    Else
        INSDBMAREPORT_QR(57) = " "
    End If
        
    If Len(Trim(txt_ttctl.Text)) > 0 Then
        INSDBMAREPORT_QR(58) = Trim(txt_ttctl.Text)
    Else
        INSDBMAREPORT_QR(58) = " "
    End If
        
    If Len(Trim(txt_lapctl.Text)) > 0 Then
        INSDBMAREPORT_QR(59) = Trim(txt_lapctl.Text)
    Else
        INSDBMAREPORT_QR(59) = " "
    End If
        
    '        If Len(Trim(cmb_mpo.Text)) > 0 Then
    '                INSDBMAREPORT_QR(60) = Trim(cmb_mpo.Text)
    '        Else
    '            INSDBMAREPORT_QR(60) = " "
    '        End If

    If Len(Trim(cmbmpo.Text)) > 0 Then
        INSDBMAREPORT_QR(60) = Trim(cmbmpo.Text)
    Else
        INSDBMAREPORT_QR(60) = " "
    End If
     
     
    If Len(Trim(txt_comments.Text)) > 0 Then
        INSDBMAREPORT_QR(61) = Trim(txt_comments.Text)
    Else
     INSDBMAREPORT_QR(61) = " "
            If txt_1.Text <> "" Then
             INSDBMAREPORT_QR(61) = Trim(txt_1.Text)
            Else
            End If
       
    End If
            
    '*** CYTOCHEMISTRY
    INSDBMAREPORT_QR.Execute
       
    '********** INSERTING INTO DHEMCBCDET FOR STORING INTERFACED VALUES OF CBC TEST *********'
    '********** ADDED BY MILIND ON 01-12-2006 *********'
    'CBCDETAIL
    '*** Inserting into DCBCREPORT
    '******** ADDDED BY CHARULATA**********
     
    '***** Add check if rec already inserted to avoid double entry 02/11/09 *****

    SEL_DHEMCPDET_QR(0) = req_chglen(cmb_reqno.Text)
    Set SEL_DHEMCPDET_RS = SEL_DHEMCPDET_QR.OpenResultset(2, 1)
    
    If SEL_DHEMCPDET_RS.EOF Then
        inshemcbc
    End If
    '******************************************************
    
         
    '      If Len(Trim(TXT_PT(0).Text)) > 0 Then
    '            INSCP_QR(0) = req_chglen(cmb_reqno.Text)
    '            INSCP_QR(1) = "PT"
    '            INSCP_QR(2) = Trim(TXT_PT(0).Text)
    '            SELCPRNG_QR(0) = "PT"
    '            Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
    '            INSCP_QR(3) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(TXT_PT(0).Text))
    '            INSCP_QR(4) = SELCPRNG_RS("control")
    '            INSCP_QR(5) = cmb_enterby.Text
    '            INSCP_QR(6) = SELCPRNG_RS("UOM")
    '            INSCP_QR(7) = SELCPRNG_RS("RANGE")
    '            If Option1.Value = True Then
    '
    '            INSCP_QR(8) = "ACC100"
    '            Else
    '            INSCP_QR(8) = "ACC200"
    '            End If
    '            'INSCP_QR(9) = ""
    '            'INSCP_QR(10) =
    '            INSCP_QR.Execute
    '
    '       End If
    '       If Len(Trim(txt_inr.Text)) > 0 Then
    '            INSCP_QR(0) = req_chglen(cmb_reqno.Text)
    '            INSCP_QR(1) = "INR"
    '            INSCP_QR(2) = Trim(txt_inr.Text)
    '            SELCPRNG_QR(0) = "INR"
    '            Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
    '            INSCP_QR(3) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(txt_inr.Text))
    '            INSCP_QR(4) = ""
    '            INSCP_QR(5) = cmb_enterby.Text
    '            INSCP_QR(6) = SELCPRNG_RS("UOM")
    '            INSCP_QR(7) = SELCPRNG_RS("RANGE")
    '            If Option1.Value = True Then
    '
    '            INSCP_QR(8) = "ACC100"
    '            Else
    '            INSCP_QR(8) = "ACC200"
    '                End If
    '            INSCP_QR.Execute
    '       End If
    '       If Len(Trim(TXT_PTTK.Text)) > 0 Then
    '            INSCP_QR(0) = req_chglen(cmb_reqno.Text)
    '            INSCP_QR(1) = "APTT"
    '            INSCP_QR(2) = Trim(TXT_PTTK.Text)
    '            SELCPRNG_QR(0) = "APTT"
    '            Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
    '            INSCP_QR(3) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(TXT_PTTK.Text))
    '            INSCP_QR(4) = ""
    '            INSCP_QR(5) = cmb_enterby.Text
    '            INSCP_QR(6) = SELCPRNG_RS("UOM")
    '            INSCP_QR(7) = SELCPRNG_RS("RANGE")
    '            If Option1.Value = True Then
    '
    '            INSCP_QR(8) = "ACC100"
    '            Else
    '            INSCP_QR(8) = "ACC200"
    '            End If
    '            INSCP_QR.Execute
    '       End If
    '       If Len(Trim(TXT_TT.Text)) > 0 Then
    '            INSCP_QR(0) = req_chglen(cmb_reqno.Text)
    '            INSCP_QR(1) = "TT"
    '            INSCP_QR(2) = Trim(TXT_TT.Text)
    '            SELCPRNG_QR(0) = "TT"
    '            Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
    '            INSCP_QR(3) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(TXT_TT.Text))
    '
    '            INSCP_QR(4) = Trim(txt_ttctl.Text) ''MANUAL CONTROL ADD
    '            INSCP_QR(5) = cmb_enterby.Text
    '            INSCP_QR(6) = SELCPRNG_RS("UOM")
    '            INSCP_QR(7) = SELCPRNG_RS("RANGE")
    '            If Option1.Value = True Then
    '
    '            INSCP_QR(8) = "ACC100"
    '            Else
    '            INSCP_QR(8) = "ACC200"
    '            End If
    '            INSCP_QR.Execute
    '       End If
    '       If Len(Trim(TXT_FIBRINOGEN.Text)) > 0 Then
    '            INSCP_QR(0) = req_chglen(cmb_reqno.Text)
    '            INSCP_QR(1) = "FIBRINOGEN"
    '            INSCP_QR(2) = Trim(TXT_FIBRINOGEN.Text)
    '            SELCPRNG_QR(0) = "FIBRINOGEN"
    '            Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
    '            INSCP_QR(3) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(TXT_FIBRINOGEN.Text))
    '
    '            INSCP_QR(4) = "" ''NO CONTROL ADD
    '            INSCP_QR(5) = cmb_enterby.Text
    '            INSCP_QR(6) = SELCPRNG_RS("UOM")
    '            INSCP_QR(7) = SELCPRNG_RS("RANGE")
    '            If Option1.Value = True Then
    '
    '            INSCP_QR(8) = "ACC100"
    '            Else
    '            INSCP_QR(8) = "ACC200"
    '            End If
    '            INSCP_QR.Execute
    '       End If
    '       If Len(Trim(TXT_FDPDMIER.Text)) > 0 Then
    '            Dim NRANGE As Integer
    '            INSCP_QR(0) = req_chglen(cmb_reqno.Text)
    '            INSCP_QR(1) = "D-DIMER"
    '            INSCP_QR(2) = Trim(TXT_FDPDMIER.Text)
    '            SELCPRNG_QR(0) = "D-DIMER"
    '            Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
    '
    '            'NRANGE
    '            INSCP_QR(3) = Mid(SELCPRNG_RS("RANGE"), 2)
    '
    '            INSCP_QR(4) = "" ''NO CONTROL ADD
    '            INSCP_QR(5) = cmb_enterby.Text
    '            INSCP_QR(6) = SELCPRNG_RS("UOM")
    '            INSCP_QR(7) = SELCPRNG_RS("RANGE")
    '            If Option1.Value = True Then
    '
    '            INSCP_QR(8) = "ACC100"
    '            Else
    '            INSCP_QR(8) = "ACC200"
    '            End If
    '            INSCP_QR.Execute
    '       End If
    '*** inserting into CBCSUBREP
    If Len(Trim(TXT_HB.Text)) > 0 Then
        INSCBCSUBREP_QR(0) = req_chglen(cmb_reqno.Text)
        INSCBCSUBREP_QR(1) = 1
        INSCBCSUBREP_QR(2) = GET_REVDATE(DATE11)
        INSCBCSUBREP_QR(3) = Trim(TXT_HB.Text)
        INSCBCSUBREP_QR(4) = Trim(TXT_RBC.Text)
        INSCBCSUBREP_QR(5) = Trim(TXT_TWBC.Text)
        INSCBCSUBREP_QR(6) = Trim(txt_platelets)
        INSCBCSUBREP_QR.Execute
    End If
    '*** INSERTING INTO CBCSUBREP
            
    updrrequisitn_qr(0) = cmb_enterby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = regcbo.Text ' cmb_finalby.Text
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = cmb_finalby.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute
    
    '***** surfacd marker in cbc insert
    If (SELDHEMCYTO_RS.RowCount > 0) Then
    If (Trim(SELDHEMCYTO_RS("smar")) = "YES" Or Trim(SELDHEMCYTO_RS("CYTO")) = "YES") Then
        inscbcsm_qr(0) = req_chglen(cmb_reqno.Text)
           
        If Len(Trim(SURFACEFORM.TXT_CD2.Text)) > 0 Then
            inscbcsm_qr(1) = SURFACEFORM.TXT_CD2.Text
        Else
            inscbcsm_qr(1) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD3.Text)) > 0 Then
            inscbcsm_qr(2) = SURFACEFORM.TXT_CD3.Text
        Else
            inscbcsm_qr(2) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CYTOCD3.Text)) > 0 Then
            inscbcsm_qr(3) = SURFACEFORM.TXT_CYTOCD3.Text
        Else
            inscbcsm_qr(3) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD4.Text)) > 0 Then
            inscbcsm_qr(4) = SURFACEFORM.TXT_CD4.Text
        Else
            inscbcsm_qr(4) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD5.Text)) > 0 Then
            inscbcsm_qr(5) = SURFACEFORM.TXT_CD5.Text
        Else
            inscbcsm_qr(5) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD7.Text)) > 0 Then ' corr
            inscbcsm_qr(6) = SURFACEFORM.TXT_CD7.Text
        Else
            inscbcsm_qr(6) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD8.Text)) > 0 Then
            inscbcsm_qr(7) = SURFACEFORM.TXT_CD8.Text
        Else
            inscbcsm_qr(7) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD10.Text)) > 0 Then
            inscbcsm_qr(8) = SURFACEFORM.TXT_CD10.Text
        Else
            inscbcsm_qr(8) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD11C.Text)) > 0 Then
            inscbcsm_qr(9) = SURFACEFORM.TXT_CD11C.Text
        Else
            inscbcsm_qr(9) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_cd13.Text)) > 0 Then
            inscbcsm_qr(10) = SURFACEFORM.txt_cd13.Text
        Else
            inscbcsm_qr(10) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_cd14.Text)) > 0 Then
            inscbcsm_qr(11) = SURFACEFORM.txt_cd14.Text
        Else
            inscbcsm_qr(11) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_CD15.Text)) > 0 Then
            inscbcsm_qr(12) = SURFACEFORM.txt_CD15.Text
        Else
            inscbcsm_qr(12) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_cd36.Text)) > 0 Then
            inscbcsm_qr(13) = SURFACEFORM.txt_cd36.Text
        Else
            inscbcsm_qr(13) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD19.Text)) > 0 Then
            inscbcsm_qr(14) = SURFACEFORM.TXT_CD19.Text
        Else
            inscbcsm_qr(14) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD20.Text)) > 0 Then
            inscbcsm_qr(15) = SURFACEFORM.TXT_CD20.Text
        Else
            inscbcsm_qr(15) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD22.Text)) > 0 Then
            inscbcsm_qr(16) = SURFACEFORM.TXT_CD22.Text
        Else
            inscbcsm_qr(16) = "  "
        End If
        If Len(Trim(SURFACEFORM.txt_cytocd221(1).Text)) > 0 Then
            inscbcsm_qr(17) = SURFACEFORM.txt_cytocd221(1).Text
        Else
            inscbcsm_qr(17) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD23.Text)) > 0 Then
            inscbcsm_qr(18) = SURFACEFORM.TXT_CD23.Text
        Else
            inscbcsm_qr(18) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD25.Text)) > 0 Then
            inscbcsm_qr(19) = SURFACEFORM.TXT_CD25.Text
        Else
            inscbcsm_qr(19) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_CD33.Text)) > 0 Then
            inscbcsm_qr(20) = SURFACEFORM.txt_CD33.Text
        Else
            inscbcsm_qr(20) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD34.Text)) > 0 Then
            inscbcsm_qr(21) = SURFACEFORM.TXT_CD34.Text
        Else
            inscbcsm_qr(21) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD38.Text)) > 0 Then
            inscbcsm_qr(22) = SURFACEFORM.TXT_CD38.Text
        Else
            inscbcsm_qr(22) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_cd41.Text)) > 0 Then
            inscbcsm_qr(23) = SURFACEFORM.txt_cd41.Text
        Else
            inscbcsm_qr(23) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD42.Text)) > 0 Then
            inscbcsm_qr(24) = SURFACEFORM.TXT_CD42.Text
        Else
            inscbcsm_qr(24) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD56.Text)) > 0 Then
            inscbcsm_qr(25) = SURFACEFORM.TXT_CD56.Text
        Else
            inscbcsm_qr(25) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD61.Text)) > 0 Then
            inscbcsm_qr(26) = SURFACEFORM.TXT_CD61.Text
        Else
            inscbcsm_qr(26) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD68.Text)) > 0 Then
            inscbcsm_qr(27) = SURFACEFORM.TXT_CD68.Text
        Else
            inscbcsm_qr(27) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_cytocd79a.Text)) > 0 Then  'corr
            inscbcsm_qr(28) = SURFACEFORM.txt_cytocd79a.Text
        Else
            inscbcsm_qr(28) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_CD79B.Text)) > 0 Then
            inscbcsm_qr(29) = SURFACEFORM.txt_CD79B.Text
        Else
            inscbcsm_qr(29) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_cd103.Text)) > 0 Then
            inscbcsm_qr(30) = SURFACEFORM.txt_cd103.Text
        Else
            inscbcsm_qr(30) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD117.Text)) > 0 Then
            inscbcsm_qr(31) = SURFACEFORM.TXT_CD117.Text
        Else
            inscbcsm_qr(31) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD138.Text)) > 0 Then
            inscbcsm_qr(32) = SURFACEFORM.TXT_CD138.Text
        Else
            inscbcsm_qr(32) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_ANTIMPO.Text)) > 0 Then
            inscbcsm_qr(33) = SURFACEFORM.TXT_ANTIMPO.Text
        Else
            inscbcsm_qr(33) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CYTOMU.Text)) > 0 Then
            inscbcsm_qr(34) = SURFACEFORM.TXT_CYTOMU.Text
        Else
            inscbcsm_qr(34) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_DBA44.Text)) > 0 Then  'corr
            inscbcsm_qr(35) = SURFACEFORM.TXT_DBA44.Text
        Else
            inscbcsm_qr(35) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_FMC7.Text)) > 0 Then
            inscbcsm_qr(36) = SURFACEFORM.TXT_FMC7.Text
        Else
            inscbcsm_qr(36) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_Glycophorina.Text)) > 0 Then
            inscbcsm_qr(37) = SURFACEFORM.txt_Glycophorina.Text
        Else
            inscbcsm_qr(37) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_HLADR.Text)) > 0 Then
            inscbcsm_qr(38) = SURFACEFORM.TXT_HLADR.Text
        Else
            inscbcsm_qr(38) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_kappa.Text)) > 0 Then
            inscbcsm_qr(39) = SURFACEFORM.txt_kappa.Text
        Else
            inscbcsm_qr(39) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_igm.Text)) > 0 Then
            inscbcsm_qr(40) = SURFACEFORM.txt_igm.Text
        Else
            inscbcsm_qr(40) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_lanbda.Text)) > 0 Then
            inscbcsm_qr(41) = SURFACEFORM.txt_lanbda.Text
        Else
            inscbcsm_qr(41) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_TDT.Text)) > 0 Then
            inscbcsm_qr(42) = SURFACEFORM.TXT_TDT.Text
        Else
            inscbcsm_qr(42) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_smig.Text)) > 0 Then
            inscbcsm_qr(43) = SURFACEFORM.txt_smig.Text
        Else
            inscbcsm_qr(43) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_clnd1.Text)) > 0 Then
            inscbcsm_qr(44) = SURFACEFORM.txt_clnd1.Text '"  "  'fld1
        Else
            inscbcsm_qr(44) = "  "  'fld2
        End If
            
        If Len(Trim(SURFACEFORM.txt_tcrab.Text)) > 0 Then
            inscbcsm_qr(45) = SURFACEFORM.txt_tcrab.Text
        Else
            inscbcsm_qr(45) = " "
        End If
            
        If Len(Trim(SURFACEFORM.txt_tcrgd.Text)) > 0 Then
            inscbcsm_qr(46) = SURFACEFORM.txt_tcrgd.Text
        Else
            inscbcsm_qr(46) = " "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_SMPCTG.Text)) > 0 Then
            inscbcsm_qr(47) = Trim(SURFACEFORM.TXT_SMPCTG.Text)
        Else
            inscbcsm_qr(47) = " "
        End If
            
        If SURFACEFORM.opt_lsp.Value = True Then
            inscbcsm_qr(48) = "Y"
        Else
            inscbcsm_qr(48) = "N"
        End If
            
        If SURFACEFORM.opt_ssp.Value = True Then
            inscbcsm_qr(49) = "Y"
        Else
            inscbcsm_qr(49) = "N"
        End If
            
        inscbcsm_qr.Execute
            
        If SURFACEFORM.opt_drak.Value = True Then
            UPDDREQDRNAME_QR(0) = "Dr. Ashok Kumar"
        ElseIf SURFACEFORM.opt_drak.Value = False Then
            UPDDREQDRNAME_QR(0) = ""
        End If
'        If SURFACEFORM.opt_mryb.Value = True Then
'            UPDDREQDRNAME_QR(1) = "Mr. Y Badrinath"
'        ElseIf SURFACEFORM.opt_mryb.Value = False Then
'            UPDDREQDRNAME_QR(1) = ""
'        End If
        'Commented and added by Shraddha on 05/05/10 *
        If SURFACEFORM.opt_mryb.Value = True Then
            UPDDREQDRNAME_QR(1) = "Mr. Y Badrinath"
        ElseIf SURFACEFORM.opt_mryb.Value = False And SURFACEFORM.opt_sitaram.Value = True Then
            UPDDREQDRNAME_QR(1) = "Mr. Sitaram"
        Else
            UPDDREQDRNAME_QR(1) = ""
        End If
        '*********************************************
        
        UPDDREQDRNAME_QR(2) = req_chglen(cmb_reqno.Text)
        UPDDREQDRNAME_QR.Execute
                
    End If '' smar = YES
     End If

End Sub
Public Sub assign()
'On Error GoTo MSG:
rdb.BeginTrans
If flagsave = 1 Then
assign_new
    '*** Inserting into  DCBCREPORT
    '*** CBC COUNTER
    '***** surfacd marker in cbc insert
End If


If flagsave = 2 And ADDENDUMFLAG = False Then
    '*** updating the DCBCREPORT
    '*** CBC COUNTER
    If Len(Trim(TXT_HB.Text)) > 0 Then
        UPDDBMAREPORT_QR(0) = Trim(TXT_HB.Text)
    Else
        UPDDBMAREPORT_QR(0) = "  "
    End If
    If Len(Trim(TXT_RBC.Text)) > 0 Then
        UPDDBMAREPORT_QR(1) = Trim(TXT_RBC.Text)
    Else
        UPDDBMAREPORT_QR(1) = "  "
    End If
    If Len(Trim(TXT_TWBC.Text)) > 0 Then
        UPDDBMAREPORT_QR(2) = Trim(TXT_TWBC.Text)
    Else
        UPDDBMAREPORT_QR(2) = "  "
    End If
    If Len(Trim(txt_platelets.Text)) > 0 Then
        UPDDBMAREPORT_QR(3) = Trim(txt_platelets.Text)
    Else
        UPDDBMAREPORT_QR(3) = "  "
    End If
    If Len(Trim(txt_neutrophils.Text)) > 0 Then
        UPDDBMAREPORT_QR(4) = Trim(txt_neutrophils.Text)
    Else
        UPDDBMAREPORT_QR(4) = "  "
    End If
    If Len(Trim(txt_lymphocytes.Text)) > 0 Then
        UPDDBMAREPORT_QR(5) = Trim(txt_lymphocytes.Text)
    Else
        UPDDBMAREPORT_QR(5) = Trim(txt_lymphocytes.Text)
    End If
    If Len(Trim(txt_monocytes.Text)) > 0 Then
        UPDDBMAREPORT_QR(6) = Trim(txt_monocytes.Text)
    Else
        UPDDBMAREPORT_QR(6) = "  "
    End If
    If Len(Trim(txt_eosinophils.Text)) > 0 Then
        UPDDBMAREPORT_QR(7) = Trim(txt_eosinophils.Text)
    Else
        UPDDBMAREPORT_QR(7) = "  "
    End If
    If Len(Trim(txt_basophils.Text)) > 0 Then
        UPDDBMAREPORT_QR(8) = Trim(txt_basophils.Text)
    Else
        UPDDBMAREPORT_QR(8) = "  "
    End If
    If Len(Trim(txt_immaturegranulocytes.Text)) > 0 Then
        UPDDBMAREPORT_QR(9) = Trim(txt_immaturegranulocytes.Text)
    Else
        UPDDBMAREPORT_QR(9) = "  "
    End If
    If Len(Trim(TXT_HCT.Text)) > 0 Then
        UPDDBMAREPORT_QR(10) = Trim(TXT_HCT.Text)
    Else
        UPDDBMAREPORT_QR(10) = "  "
    End If
    If Len(Trim(TXT_MCH.Text)) > 0 Then
        UPDDBMAREPORT_QR(11) = Trim(TXT_MCH.Text)
    Else
        UPDDBMAREPORT_QR(11) = "  "
    End If
    If Len(Trim(TXT_MCV.Text)) > 0 Then
        UPDDBMAREPORT_QR(12) = Trim(TXT_MCV.Text)
    Else
        UPDDBMAREPORT_QR(12) = "  "
    End If
    If Len(Trim(TXT_MCHC.Text)) > 0 Then
        UPDDBMAREPORT_QR(13) = Trim(TXT_MCHC.Text)
    Else
        UPDDBMAREPORT_QR(13) = "  "
    End If
    If Len(Trim(TXT_MPV.Text)) > 0 Then
        UPDDBMAREPORT_QR(14) = Trim(TXT_MPV.Text)
    Else
        UPDDBMAREPORT_QR(14) = "  "
    End If
    If Len(Trim(TXT_RDW.Text)) > 0 Then
        UPDDBMAREPORT_QR(15) = Trim(TXT_RDW.Text)
    Else
        UPDDBMAREPORT_QR(15) = "  "
    End If
    If Len(Trim(TXT_ABNORMAL.Text)) > 0 Then
        UPDDBMAREPORT_QR(16) = Trim(TXT_ABNORMAL.Text)
    Else
        UPDDBMAREPORT_QR(16) = "  "
    End If
    '*** CBC COUNTER
    '*** CBC MANUAL
    If Len(Trim(txt_mneutrophils.Text)) > 0 Then
        UPDDBMAREPORT_QR(17) = Trim(txt_mneutrophils.Text)
    Else
        UPDDBMAREPORT_QR(17) = "  "
    End If
    If Len(Trim(txt_mlymphocytes.Text)) > 0 Then
        UPDDBMAREPORT_QR(18) = Trim(txt_mlymphocytes.Text)
    Else
        UPDDBMAREPORT_QR(18) = "  "
    End If
    '****** 16/10/08 Shraddha Pandit *****
    '                If Len(Trim(txt_prolympho.Text)) > 0 Then
    '                    UPDDBMAREPORT_QR(62) = Trim(txt_prolympho.Text)
    '                Else
    '                    UPDDBMAREPORT_QR(62) = "  "
    '                End If
    '*************************************
    If Len(Trim(txt_mmonocytes.Text)) > 0 Then
        UPDDBMAREPORT_QR(19) = Trim(txt_mmonocytes.Text)
    Else
        UPDDBMAREPORT_QR(19) = "  "
    End If
    If Len(Trim(txt_meosinophils.Text)) > 0 Then
        UPDDBMAREPORT_QR(20) = Trim(txt_meosinophils.Text)
    Else
        UPDDBMAREPORT_QR(20) = "  "
    End If
    If Len(Trim(txt_mbasophills.Text)) > 0 Then
        UPDDBMAREPORT_QR(21) = Trim(txt_mbasophills.Text)
    Else
        UPDDBMAREPORT_QR(21) = "  "
    End If
    If Len(Trim(txt_blasts.Text)) > 0 Then
        UPDDBMAREPORT_QR(22) = Trim(txt_blasts.Text)
    Else
        UPDDBMAREPORT_QR(22) = "  "
    End If
'    If Len(Trim(TXT_PROMYELOCYTESWBC.Text)) > 0 Then
'        UPDDBMAREPORT_QR(23) = Trim(TXT_PROMYELOCYTESWBC.Text)
'    Else
'        UPDDBMAREPORT_QR(23) = "  "
'    End If
    If Len(Trim(txt_myelocytes.Text)) > 0 Then
        UPDDBMAREPORT_QR(23) = Trim(txt_myelocytes.Text)
    Else
        UPDDBMAREPORT_QR(23) = "  "
    End If
    If Len(Trim(txt_metamyelocytes.Text)) > 0 Then
        UPDDBMAREPORT_QR(24) = Trim(txt_metamyelocytes.Text)
    Else
        UPDDBMAREPORT_QR(24) = "  "
    End If
    If Len(Trim(txt_plasmacells.Text)) > 0 Then
        UPDDBMAREPORT_QR(25) = Trim(txt_plasmacells.Text)
    Else
        UPDDBMAREPORT_QR(25) = "  "
    End If
    If Len(Trim(txt_alymphocytes.Text)) > 0 Then
        UPDDBMAREPORT_QR(26) = Trim(txt_alymphocytes.Text)
    Else
        UPDDBMAREPORT_QR(26) = "  "
    End If
    If Len(Trim(txt_prolympho.Text)) > 0 Then
        UPDDBMAREPORT_QR(27) = Trim(txt_prolympho.Text)
    Else
        UPDDBMAREPORT_QR(27) = "  "
    End If
    If Len(Trim(txt_Sezarycells.Text)) > 0 Then
        UPDDBMAREPORT_QR(28) = Trim(txt_Sezarycells.Text)
    Else
        UPDDBMAREPORT_QR(28) = "  "
    End If
    If Len(Trim(txt_hairycells.Text)) > 0 Then
        UPDDBMAREPORT_QR(29) = Trim(txt_hairycells.Text)
    Else
        UPDDBMAREPORT_QR(29) = "  "
    End If
    If Len(Trim(txt_othercells.Text)) > 0 Then
        UPDDBMAREPORT_QR(30) = Trim(txt_othercells.Text)
    Else
        UPDDBMAREPORT_QR(30) = "  "
    End If
    If Len(Trim(TXT_NORMOBLASTS.Text)) > 0 Then
        UPDDBMAREPORT_QR(31) = Trim(TXT_NORMOBLASTS.Text)
    Else
        UPDDBMAREPORT_QR(31) = "  "
    End If
    If Len(Trim(TXT_OTHERCELLSRBC.Text)) > 0 Then
        UPDDBMAREPORT_QR(32) = Trim(TXT_OTHERCELLSRBC.Text)
    Else
        UPDDBMAREPORT_QR(32) = "  "
    End If
    If Len(Trim(txt_rbcmorphology.Text)) > 0 Then
        UPDDBMAREPORT_QR(33) = Trim(txt_rbcmorphology.Text)
    Else
        UPDDBMAREPORT_QR(33) = "  "
    End If
    '********* Shraddha Pandit 17/10/08 ********
    '                If Len(Trim(TXT_LP.Text)) > 0 Then
    '                    UPDDBMAREPORT_QR(34) = Trim(TXT_LP.Text)
    '                Else
    '                    UPDDBMAREPORT_QR(34) = "  "
    '                End If
    '******************************************
    If Len(Trim(txt_reticcount.Text)) > 0 Then
        UPDDBMAREPORT_QR(34) = Trim(txt_reticcount.Text)
    Else
        UPDDBMAREPORT_QR(34) = "  "
    End If
    If Len(Trim(txt_malarialparasite.Text)) > 0 Then
        UPDDBMAREPORT_QR(35) = " " 'Trim(txt_malarialparasite.Text)
    Else
        UPDDBMAREPORT_QR(35) = "  "
    End If
                
    If Len(Trim(cmbmpo.Text)) > 0 Then
        UPDDBMAREPORT_QR(47) = Trim(cmbmpo.Text)
    Else
        UPDDBMAREPORT_QR(47) = "  "
    End If

           
    DCBCREPA
                
                
    If Len(Trim(txt_esr.Text)) > 0 Then '18/11/08
        UPDDBMAREPORT_QR(36) = Trim(txt_esr.Text)
    Else
        UPDDBMAREPORT_QR(36) = "  "
    End If
    'UPDDBMAREPORT_QR(37) = "  "

    'If Len(Trim(TXT_MORPHOLOGY.Text)) > 0 Then
    
    If Len(Trim(TXT_PROMYELOCYTESWBC.Text)) > 0 Then 'Shraddha 06/05/20
        'UPDDBMAREPORT_QR(38) = Trim(TXT_MORPHOLOGY.Text)
        'UPDDBMAREPORT_QR(38) = Trim(CMB_MORPHO.Text)
        UPDDBMAREPORT_QR(37) = Trim(TXT_PROMYELOCYTESWBC.Text) 'Shraddha 13/11/08
        'UPDDBMAREPORT_QR(38) = Trim(CMB_MORPHO.Text) 'Shraddha 13/04/09
    Else
        UPDDBMAREPORT_QR(37) = "  "
    End If
    '*** CBC MANUAL
    '*** COAGULATION PROFILE
    If Len(Trim(txt_inr.Text)) > 0 Then
        UPDDBMAREPORT_QR(38) = Trim(txt_inr.Text)
    Else
        UPDDBMAREPORT_QR(38) = "  "
    End If
    If Len(Trim(TXT_PT(0).Text)) > 0 Then
        UPDDBMAREPORT_QR(39) = Trim(TXT_PT(0).Text)
    Else
        UPDDBMAREPORT_QR(39) = "  "
    End If
    If Len(Trim(txt_ptctl.Text)) > 0 Then
        UPDDBMAREPORT_QR(40) = Trim(txt_ptctl.Text) 'platelet1
    Else
        UPDDBMAREPORT_QR(40) = "  "
    End If
    If Len(Trim(TXT_PTTK.Text)) > 0 Then
        UPDDBMAREPORT_QR(41) = Trim(TXT_PTTK.Text)
    Else
        UPDDBMAREPORT_QR(41) = "  "
    End If
    If Len(Trim(TXT_TT.Text)) > 0 Then
        UPDDBMAREPORT_QR(42) = Trim(TXT_TT.Text)
    Else
        UPDDBMAREPORT_QR(42) = "  "
    End If
    If Len(Trim(TXT_FIBRINOGEN.Text)) > 0 Then
        UPDDBMAREPORT_QR(43) = Trim(TXT_FIBRINOGEN.Text)
    Else
        UPDDBMAREPORT_QR(43) = "  "
    End If
    If Len(Trim(TXT_FDPDMIER.Text)) > 0 Then
        UPDDBMAREPORT_QR(44) = Trim(TXT_FDPDMIER.Text)
    Else
        UPDDBMAREPORT_QR(44) = "  "
    End If
    If Len(Trim(TXT_OTHERS.Text)) > 0 Then
        UPDDBMAREPORT_QR(45) = Trim(TXT_OTHERS.Text)
    Else
        UPDDBMAREPORT_QR(45) = "  "
    End If
    '*** COAGULATION PROFILE
    '*** CYTOCHEMISTRY
    If Len(Trim(txt_lapscore.Text)) > 0 Then
        UPDDBMAREPORT_QR(46) = Trim(txt_lapscore.Text)
    Else
        UPDDBMAREPORT_QR(46) = "  "
    End If
    '                If Len(Trim(TXT_MPO.Text)) > 0 Then
    '                    UPDDBMAREPORT_QR(48) = Trim(TXT_MPO.Text)
    '                Else
    '                    UPDDBMAREPORT_QR(48) = "  "
    '                End If
    If Len(Trim(cmb_nse.Text)) > 0 Then
        UPDDBMAREPORT_QR(48) = Trim(cmb_nse.Text)
    Else
        UPDDBMAREPORT_QR(48) = "  "
    End If
'    If Len(Trim(TXT_PAS.Text)) > 0 Then
'        UPDDBMAREPORT_QR(50) = Trim(TXT_PAS.Text)
'    Else
'        UPDDBMAREPORT_QR(50) = "  "
'    End If
    'PAS is of no use. So used to store value of promonocyte
    '************ Added by Shraddha Pandit on 06/04/09 *************
    If Len(Trim(txt_promonocyte.Text)) > 0 Then
        UPDDBMAREPORT_QR(49) = Trim(txt_promonocyte.Text)
    Else
        UPDDBMAREPORT_QR(49) = "  "
    End If
'***************************************************************
    
    If Len(Trim(TXT_NSE.Text)) > 0 Then
        UPDDBMAREPORT_QR(50) = Trim(TXT_NSE.Text)
    Else
        UPDDBMAREPORT_QR(50) = "  "
    End If
    If Len(Trim(TXT_TRAP.Text)) > 0 Then
        UPDDBMAREPORT_QR(51) = Trim(TXT_TRAP.Text)
    Else
        UPDDBMAREPORT_QR(51) = "  "
    End If
                
    If Len(Trim(TXT_ADEQUACY.Text)) > 0 Then
        UPDDBMAREPORT_QR(52) = Trim(TXT_ADEQUACY.Text)
    Else
        UPDDBMAREPORT_QR(52) = "  "
    End If
    '******** Shraddha Pandit 17/10/08 *******
    If Len(Trim(TXT_LP.Text)) > 0 Then
        UPDDBMAREPORT_QR(52) = Trim(TXT_LP.Text)
    Else
        UPDDBMAREPORT_QR(52) = "  "
    End If
    '*****************************************
    '*** CYTOCHEMISTRY
    If Len(Trim(txt_impression.Text)) > 0 Then
        'UPDDBMAREPORT_QR(53) = "  "
        UPDDBMAREPORT_QR(53) = Trim(txt_impression.Text)
    Else
        UPDDBMAREPORT_QR(53) = "  "
    End If
                
    If Len(Trim(Text1.Text)) > 0 Then
        UPDDBMAREPORT_QR(54) = Trim(Text1.Text) ' platelet morphology
    Else
        UPDDBMAREPORT_QR(54) = "  "
    End If
                
    '****** 13/11/08 Shraddha Pandit *****
    If Len(Trim(txt_prolympho.Text)) > 0 Then
        UPDDBMAREPORT_QR(54) = Trim(CMB_MORPHO.Text)
    Else
        UPDDBMAREPORT_QR(54) = "  "
    End If
    '*************************************

    If Len(Trim(txt_apttctl.Text)) > 0 Then
        UPDDBMAREPORT_QR(55) = Trim(txt_apttctl.Text)
    Else
        UPDDBMAREPORT_QR(55) = " "
    End If
    If Len(Trim(txt_ttctl.Text)) > 0 Then
        UPDDBMAREPORT_QR(56) = Trim(txt_ttctl.Text)
    Else
        UPDDBMAREPORT_QR(56) = " "
    End If
    If Len(Trim(txt_lapctl.Text)) > 0 Then
        UPDDBMAREPORT_QR(57) = Trim(txt_lapctl.Text)
    Else
        UPDDBMAREPORT_QR(57) = " "
    End If
                
    '                If Len(Trim(cmb_mpo.Text)) > 0 Then
    '                    UPDDBMAREPORT_QR(59) = Trim(cmb_mpo.Text)
    '                Else
    '                    UPDDBMAREPORT_QR(59) = " "
    '                End If

    If Len(Trim(cmbmpo.Text)) > 0 Then
        UPDDBMAREPORT_QR(58) = Trim(cmbmpo.Text)
    Else
        UPDDBMAREPORT_QR(58) = " "
    End If
        
    If Len(Trim(txt_comments.Text)) > 0 Then
        UPDDBMAREPORT_QR(59) = Trim(txt_comments.Text)
    Else
        UPDDBMAREPORT_QR(59) = " "
            If txt_1.Text <> "" Then
             UPDDBMAREPORT_QR(59) = Trim(txt_1.Text)
            Else
            End If
        
    End If
    
    
                             
    UPDDBMAREPORT_QR(60) = req_chglen(cmb_reqno.Text)
    UPDDBMAREPORT_QR.Execute
                
    'FOR UPDATING CBCDETAILS
    CBCDTLUPDT
    
    ''************ FOR COAGULATION PROFILE
    For I = 0 To 5
        If I = 0 Then
            If Len(Trim(TXT_PT(0).Text)) > 0 Then
                UPDCPDET_QR(0) = Trim(TXT_PT(0).Text)
                SELCPRNG_QR(0) = "PT"
                Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
                UPDCPDET_QR(1) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(TXT_PT(0).Text))
                UPDCPDET_QR(2) = Trim(txt_ptctl.Text)
            Else
                UPDCPDET_QR(0) = ""
                UPDCPDET_QR(1) = ""
                UPDCPDET_QR(2) = ""
            End If
            
            UPDCPDET_QR(3) = cmb_enterby.Text
'
'            If Option1.Value = True Then
'                UPDCPDET_QR(4) = "ACL 100"
'            ElseIf Option2.Value = True Then
'                UPDCPDET_QR(4) = "ACL 200"
'            ElseIf Option3.Value = True Then
'                UPDCPDET_QR(4) = "ELP 001"
'            Else
                UPDCPDET_QR(4) = cmb_machname.Text
'            End If
            
            UPDCPDET_QR(5) = req_chglen(cmb_reqno.Text)
            UPDCPDET_QR(6) = "PT"
            Set UPDCPDET_RS = UPDCPDET_QR.OpenResultset(2, 1)
        End If
        
        If I = 1 Then
            If Len(Trim(txt_inr.Text)) > 0 Then
                UPDCPDET_QR(0) = Trim(txt_inr.Text)
                SELCPRNG_QR(0) = "INR"
                Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
                UPDCPDET_QR(1) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(txt_inr.Text))
            Else
                UPDCPDET_QR(0) = ""
                UPDCPDET_QR(1) = ""
            End If
            
            UPDCPDET_QR(2) = ""
            UPDCPDET_QR(3) = cmb_enterby.Text
'
'            If Option1.Value = True Then
'                UPDCPDET_QR(4) = "ACL 100"
'            ElseIf Option2.Value = True Then
'                UPDCPDET_QR(4) = "ACL 200"
'            ElseIf Option3.Value = True Then
'                UPDCPDET_QR(4) = "ELP 001"
'            Else
                UPDCPDET_QR(4) = cmb_machname.Text
'            End If
'
            UPDCPDET_QR(5) = req_chglen(cmb_reqno.Text)
            UPDCPDET_QR(6) = "INR"
            UPDCPDET_QR.Execute
        End If
        
        If I = 2 Then
            If Len(Trim(TXT_PTTK.Text)) > 0 Then
                UPDCPDET_QR(0) = Trim(TXT_PTTK.Text)
                SELCPRNG_QR(0) = "APTT"
                Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
                UPDCPDET_QR(1) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(TXT_PTTK.Text))
                UPDCPDET_QR(2) = Trim(txt_apttctl.Text)
            Else
                UPDCPDET_QR(0) = ""
                UPDCPDET_QR(1) = ""
                UPDCPDET_QR(2) = ""
            End If
            
            UPDCPDET_QR(3) = cmb_enterby.Text
            
'            If Option1.Value = True Then
'                UPDCPDET_QR(4) = "ACL 100"
'            ElseIf Option2.Value = True Then
'                UPDCPDET_QR(4) = "ACL 200"
'            ElseIf Option1.Value = True Then
'                UPDCPDET_QR(4) = "ELP 001"
'            Else
                UPDCPDET_QR(4) = cmb_machname.Text
'            End If
            
            UPDCPDET_QR(5) = req_chglen(cmb_reqno.Text)
            UPDCPDET_QR(6) = "APTT"
            Set UPDCPDET_RS = UPDCPDET_QR.OpenResultset(2, 1)
        End If
        
        If I = 3 Then
            If Len(Trim(TXT_TT.Text)) > 0 Then
                UPDCPDET_QR(0) = Trim(TXT_TT.Text)
                SELCPRNG_QR(0) = "TT"
                Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
                UPDCPDET_QR(1) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(TXT_TT.Text))
                UPDCPDET_QR(2) = Trim(txt_ttctl.Text)
            Else
                UPDCPDET_QR(0) = ""
                UPDCPDET_QR(1) = ""
                UPDCPDET_QR(2) = ""
            End If
            
            UPDCPDET_QR(3) = cmb_enterby.Text
'
'            If Option1.Value = True Then
'                UPDCPDET_QR(4) = "ACL 100"
'            ElseIf Option2.Value = True Then
'                UPDCPDET_QR(4) = "ACL 200"
'            ElseIf Option3.Value = True Then
'                UPDCPDET_QR(4) = "ELP 01"
'            Else
                UPDCPDET_QR(4) = cmb_machname.Text
'            End If
            
            UPDCPDET_QR(5) = req_chglen(cmb_reqno.Text)
            UPDCPDET_QR(6) = "TT"
            Set UPDCPDET_RS = UPDCPDET_QR.OpenResultset(2, 1)
        End If
        
        If I = 4 Then
            If Len(Trim(TXT_FIBRINOGEN.Text)) > 0 Then
                UPDCPDET_QR(0) = Trim(TXT_FIBRINOGEN.Text)
                SELCPRNG_QR(0) = "FIBRINOGEN"
                Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
                UPDCPDET_QR(1) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(TXT_FIBRINOGEN.Text))
            Else
                UPDCPDET_QR(0) = ""
                UPDCPDET_QR(1) = ""
            End If
            
            UPDCPDET_QR(2) = ""
            UPDCPDET_QR(3) = cmb_enterby.Text
            
'            If Option1.Value = True Then
'                UPDCPDET_QR(4) = "ACL 100"
'            ElseIf Option2.Value = True Then
'                UPDCPDET_QR(4) = "ACL 200"
'            ElseIf Option3.Value = True Then
'                UPDCPDET_QR(4) = "ELP 001"
'            Else
                UPDCPDET_QR(4) = cmb_machname.Text
'            End If
'
            UPDCPDET_QR(5) = req_chglen(cmb_reqno.Text)
            UPDCPDET_QR(6) = "FIBRINOGEN"
            Set UPDCPDET_RS = UPDCPDET_QR.OpenResultset(2, 1)
        End If
        
        If I = 5 Then
            If Len(Trim(TXT_FDPDMIER.Text)) > 0 Then
                Dim DDIMER As Integer
                Dim NRANGE As String
                UPDCPDET_QR(0) = Trim(TXT_FDPDMIER.Text)
                SELCPRNG_QR(0) = "D-DIMER"
                Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
                DDIMER = Val(Mid(Trim(SELCPRNG_RS("RANGE")), 2, 3))
                'ddimer = 200
                NRANGE = Trim(TXT_FDPDMIER.Text)
                
                If InStr(1, NRANGE, "<") > 0 Then
                    Do While InStr(NRANGE, "<") > 0
                        NRANGE = Replace(NRANGE, "<", "*")
                    Loop
                End If
                If InStr(1, NRANGE, ">") > 0 Then
                    Do While InStr(NRANGE, ">") > 0
                        NRANGE = Replace(NRANGE, ">", "*")
                    Loop
                End If
                                    
                If Val(Mid(NRANGE, 2, 3)) < DDIMER Then
                    UPDCPDET_QR(1) = "L"
                ElseIf Val(Mid(NRANGE, 2, 3)) > DDIMER Then
                    UPDCPDET_QR(1) = "H"
                ElseIf Val(Mid(NRANGE, 2, 3)) = DDIMER Then
                    UPDCPDET_QR(1) = "N"
                End If
                
            Else
                UPDCPDET_QR(0) = ""
                UPDCPDET_QR(1) = ""
            End If
            
            UPDCPDET_QR(2) = ""
            UPDCPDET_QR(3) = cmb_enterby.Text
'
'            If Option1.Value = True Then
'                UPDCPDET_QR(4) = "ACL 100"
'            ElseIf Option2.Value = True Then
'                UPDCPDET_QR(4) = "ACL 200"
'            ElseIf Option3.Value = True Then
'                UPDCPDET_QR(4) = "ELP 01"
'            Else
                UPDCPDET_QR(4) = cmb_machname.Text
'            End If
            
            UPDCPDET_QR(5) = req_chglen(cmb_reqno.Text)
            UPDCPDET_QR(6) = "D-DIMER"
            Set UPDCPDET_RS = UPDCPDET_QR.OpenResultset(2, 1)
        End If
    Next
                   
    If diffcountflag = True Then
        If Len(Trim(TXT_HB.Text)) > 0 Then
            If SELCBCSUBREPFLAG = False Then
                SELCBCSUBREP_QR(0) = req_chglen(cmb_reqno.Text)
                Set SELCBCSUBREP_RS = SELCBCSUBREP_QR.OpenResultset(1, 3)
                SELCBCSUBREPFLAG = True
            Else
                SELCBCSUBREP_QR(0) = req_chglen(cmb_reqno.Text)
                SELCBCSUBREP_RS.MoreResults
                SELCBCSUBREP_RS.Requery
            End If
            INSCBCSUBREP_QR(0) = req_chglen(cmb_reqno.Text)
            If SELCBCSUBREP_RS.RowCount > 0 Then
                INSCBCSUBREP_QR(1) = SELCBCSUBREP_RS.RowCount + 1
            Else
                INSCBCSUBREP_QR(1) = 1
            End If
            INSCBCSUBREP_QR(2) = GET_REVDATE(DATE11)
            INSCBCSUBREP_QR(3) = Trim(TXT_HB.Text)
            INSCBCSUBREP_QR(4) = Trim(TXT_RBC.Text)
            INSCBCSUBREP_QR(5) = Trim(TXT_TWBC.Text)
            INSCBCSUBREP_QR(6) = Trim(txt_platelets)
            INSCBCSUBREP_QR.Execute
        End If
    End If
    '*** updating the DCBCREPORT
    updrrequisitn_qr(0) = cmb_enterby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = regcbo.Text ' cmb_finalby.Text
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = cmb_finalby.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
    updrrequisitn_qr.Execute
    '' UPDATE SURFACE MARKER IN CBC HL
    If SELDHEMCYTO_RS.EOF = False Then
    If Trim(SELDHEMCYTO_RS("smar")) = "YES" Or Trim(SELDHEMCYTO_RS("CYTO")) = "YES" Then
        If Len(Trim(SURFACEFORM.TXT_CD2.Text)) > 0 Then
            updcbcsm_qr(0) = SURFACEFORM.TXT_CD2.Text
        Else
            updcbcsm_qr(0) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD3.Text)) > 0 Then
            updcbcsm_qr(1) = SURFACEFORM.TXT_CD3.Text
        Else
            updcbcsm_qr(1) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CYTOCD3.Text)) > 0 Then
            updcbcsm_qr(2) = SURFACEFORM.TXT_CYTOCD3.Text
        Else
            updcbcsm_qr(2) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD4.Text)) > 0 Then
            updcbcsm_qr(3) = SURFACEFORM.TXT_CD4.Text
        Else
            updcbcsm_qr(3) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD5.Text)) > 0 Then
            updcbcsm_qr(4) = SURFACEFORM.TXT_CD5.Text
        Else
            updcbcsm_qr(4) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD7.Text)) > 0 Then ' corr
            updcbcsm_qr(5) = SURFACEFORM.TXT_CD7.Text
        Else
            updcbcsm_qr(5) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD8.Text)) > 0 Then
            updcbcsm_qr(6) = SURFACEFORM.TXT_CD8.Text
        Else
            updcbcsm_qr(6) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD10.Text)) > 0 Then
            updcbcsm_qr(7) = SURFACEFORM.TXT_CD10.Text
        Else
            updcbcsm_qr(7) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD11C.Text)) > 0 Then
            updcbcsm_qr(8) = SURFACEFORM.TXT_CD11C.Text
        Else
            updcbcsm_qr(8) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_cd13.Text)) > 0 Then
            updcbcsm_qr(9) = SURFACEFORM.txt_cd13.Text
        Else
            updcbcsm_qr(9) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_cd14.Text)) > 0 Then
            updcbcsm_qr(10) = SURFACEFORM.txt_cd14.Text
        Else
            updcbcsm_qr(10) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_CD15.Text)) > 0 Then
            updcbcsm_qr(11) = SURFACEFORM.txt_CD15.Text
        Else
            updcbcsm_qr(11) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_cd36.Text)) > 0 Then
            updcbcsm_qr(12) = SURFACEFORM.txt_cd36.Text
        Else
            updcbcsm_qr(12) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD19.Text)) > 0 Then
            updcbcsm_qr(13) = SURFACEFORM.TXT_CD19.Text
        Else
            updcbcsm_qr(13) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD20.Text)) > 0 Then
            updcbcsm_qr(14) = SURFACEFORM.TXT_CD20.Text
        Else
            updcbcsm_qr(14) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD22.Text)) > 0 Then
            updcbcsm_qr(15) = SURFACEFORM.TXT_CD22.Text
        Else
            updcbcsm_qr(15) = "  "
        End If
        If Len(Trim(SURFACEFORM.txt_cytocd221(1).Text)) > 0 Then
            updcbcsm_qr(16) = SURFACEFORM.txt_cytocd221(1).Text
        Else
            updcbcsm_qr(16) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD23.Text)) > 0 Then
            updcbcsm_qr(17) = SURFACEFORM.TXT_CD23.Text
        Else
            updcbcsm_qr(17) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD25.Text)) > 0 Then
            updcbcsm_qr(18) = SURFACEFORM.TXT_CD25.Text
        Else
            updcbcsm_qr(18) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_CD33.Text)) > 0 Then
            updcbcsm_qr(19) = SURFACEFORM.txt_CD33.Text
        Else
            updcbcsm_qr(19) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD34.Text)) > 0 Then
            updcbcsm_qr(20) = SURFACEFORM.TXT_CD34.Text
        Else
            updcbcsm_qr(20) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD38.Text)) > 0 Then
            updcbcsm_qr(21) = SURFACEFORM.TXT_CD38.Text
        Else
            updcbcsm_qr(21) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_cd41.Text)) > 0 Then
            updcbcsm_qr(22) = SURFACEFORM.txt_cd41.Text
        Else
            updcbcsm_qr(22) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD42.Text)) > 0 Then
            updcbcsm_qr(23) = SURFACEFORM.TXT_CD42.Text
        Else
            updcbcsm_qr(23) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD56.Text)) > 0 Then
            updcbcsm_qr(24) = SURFACEFORM.TXT_CD56.Text
        Else
            updcbcsm_qr(24) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD61.Text)) > 0 Then
            updcbcsm_qr(25) = SURFACEFORM.TXT_CD61.Text
        Else
            updcbcsm_qr(25) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD68.Text)) > 0 Then
            updcbcsm_qr(26) = SURFACEFORM.TXT_CD68.Text
        Else
            updcbcsm_qr(26) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_cytocd79a.Text)) > 0 Then  'corr
            updcbcsm_qr(27) = SURFACEFORM.txt_cytocd79a.Text
        Else
            updcbcsm_qr(27) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_CD79B.Text)) > 0 Then
            updcbcsm_qr(28) = SURFACEFORM.txt_CD79B.Text
        Else
            updcbcsm_qr(28) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_cd103.Text)) > 0 Then
            updcbcsm_qr(29) = SURFACEFORM.txt_cd103.Text
        Else
            updcbcsm_qr(29) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD117.Text)) > 0 Then
            updcbcsm_qr(30) = SURFACEFORM.TXT_CD117.Text
        Else
            updcbcsm_qr(30) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CD138.Text)) > 0 Then
            updcbcsm_qr(31) = SURFACEFORM.TXT_CD138.Text
        Else
            updcbcsm_qr(31) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_ANTIMPO.Text)) > 0 Then
            updcbcsm_qr(32) = SURFACEFORM.TXT_ANTIMPO.Text
        Else
            updcbcsm_qr(32) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_CYTOMU.Text)) > 0 Then
            updcbcsm_qr(33) = SURFACEFORM.TXT_CYTOMU.Text
        Else
            updcbcsm_qr(33) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_DBA44.Text)) > 0 Then  'corr
            updcbcsm_qr(34) = SURFACEFORM.TXT_DBA44.Text
        Else
            updcbcsm_qr(34) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_FMC7.Text)) > 0 Then
            updcbcsm_qr(35) = SURFACEFORM.TXT_FMC7.Text
        Else
            updcbcsm_qr(35) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_Glycophorina.Text)) > 0 Then
            updcbcsm_qr(36) = SURFACEFORM.txt_Glycophorina.Text
        Else
            updcbcsm_qr(36) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_HLADR.Text)) > 0 Then
            updcbcsm_qr(37) = SURFACEFORM.TXT_HLADR.Text
        Else
            updcbcsm_qr(37) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_kappa.Text)) > 0 Then
            updcbcsm_qr(38) = SURFACEFORM.txt_kappa.Text
        Else
            updcbcsm_qr(38) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_igm.Text)) > 0 Then
            updcbcsm_qr(39) = SURFACEFORM.txt_igm.Text
        Else
            updcbcsm_qr(39) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_lanbda.Text)) > 0 Then
            updcbcsm_qr(40) = SURFACEFORM.txt_lanbda.Text
        Else
            updcbcsm_qr(40) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_TDT.Text)) > 0 Then
            updcbcsm_qr(41) = SURFACEFORM.TXT_TDT.Text
        Else
            updcbcsm_qr(41) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_smig.Text)) > 0 Then
            updcbcsm_qr(42) = SURFACEFORM.txt_smig.Text
        Else
            updcbcsm_qr(42) = "  "
        End If
            
        If Len(Trim(SURFACEFORM.txt_clnd1.Text)) > 0 Then
            updcbcsm_qr(43) = SURFACEFORM.txt_clnd1.Text
        Else
            updcbcsm_qr(43) = " "
        End If
            
        If Len(Trim(SURFACEFORM.txt_tcrab.Text)) > 0 Then
            updcbcsm_qr(44) = SURFACEFORM.txt_tcrab.Text
        Else
            updcbcsm_qr(44) = " "
        End If
            
        If Len(Trim(SURFACEFORM.txt_tcrgd.Text)) > 0 Then
            updcbcsm_qr(45) = SURFACEFORM.txt_tcrgd.Text
        Else
            updcbcsm_qr(45) = " "
        End If
            
        If Len(Trim(SURFACEFORM.TXT_SMPCTG.Text)) > 0 Then
            updcbcsm_qr(46) = Trim(SURFACEFORM.TXT_SMPCTG.Text)
        Else
            updcbcsm_qr(46) = " "
        End If
        
        If SURFACEFORM.opt_lsp.Value = True Then
            updcbcsm_qr(47) = "Y"
        Else
            updcbcsm_qr(47) = "N"
        End If
            
        If SURFACEFORM.opt_ssp.Value = True Then
            updcbcsm_qr(48) = "Y"
        Else
            updcbcsm_qr(48) = "N"
        End If
                  
        updcbcsm_qr(49) = req_chglen(cmb_reqno.Text)
                    updcbcsm_qr.Execute
            
        If SURFACEFORM.opt_drak.Value = True Then
            UPDDREQDRNAME_QR(0) = "Dr. Ashok Kumar"
        ElseIf SURFACEFORM.opt_drak.Value = False Then
            UPDDREQDRNAME_QR(0) = ""
        End If
'        If SURFACEFORM.opt_mryb.Value = True Then
'            UPDDREQDRNAME_QR(1) = "Mr. Y Badrinath"
'        ElseIf SURFACEFORM.opt_mryb.Value = False Then
'            UPDDREQDRNAME_QR(1) = ""
'        End If
'***** Commented and added by Shraddha on 05/05/10 *****
        If SURFACEFORM.opt_mryb.Value = True Then
            UPDDREQDRNAME_QR(1) = "Mr. Y Badrinath"
        ElseIf SURFACEFORM.opt_mryb.Value = False And SURFACEFORM.opt_sitaram.Value = True Then
            UPDDREQDRNAME_QR(1) = "Mr. Sitaram"
        Else
            UPDDREQDRNAME_QR(1) = ""
        End If
        '*****************************************
        UPDDREQDRNAME_QR(2) = req_chglen(cmb_reqno.Text)
        UPDDREQDRNAME_QR.Execute
    End If '' smar = YES
        End If
    '' UPDATE SURFACE MARKER IN CBC HL ENDS
    ''******ADDENDUM*******************************
    SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
    SELDADDENDUM_RS.MoreResults
    SELDADDENDUM_RS.Requery
    
    'ADDED BY MILIND ON 16-08-2006 FOR REMOVING DEFAULT ADDENDUM REPORT GENERATION
    Dim cmt_qr As rdoQuery
    Dim cmt_rs As rdoResultset
    str = "select commitflag from rislib.rrequisitn where reqnno=?"
    Set cmt_qr = rdb.CreateQuery("cmt_qr", str)
    cmt_qr(0) = req_chglen(cmb_reqno.Text)
    Set cmt_rs = cmt_qr.OpenResultset(2, 3)
    
    If SELDADDENDUM_RS.RowCount > 0 Then
        SELDADDENDUM_RS.MoveLast
        INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
        INSDADDENDUM_QR(1) = SELDADDENDUM_RS("ADDEND_NO") + 1
        INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        INSDADDENDUM_QR(3) = time1
        INSDADDENDUM_QR(4) = cmb_addendumby.Text
        INSDADDENDUM_QR(5) = Null
        INSDADDENDUM_QR(5) = txt_addendum.Text
        INSDADDENDUM_QR.Execute
    Else 'ADDED BY MILIND ON 16-08-2006
        If cmt_rs(0) = "Y" Then
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
    ''******ADDENDUM*******************************
MsgBox "Record is Saved"
End If

'*************** Shraddha Pandit 30/03/09 *****************
'If ADDENDUMFLAG = True then 'commented and added on 01/04/09
If ADDENDUMFLAG = True Then 'And SSTab1.Caption = "Suface Marker" Then
    If Trim(cmb_addendumby.Text) = "" Then
        
        MsgBox "The Addendum By field is EMPTY."
        If cmb_addendumby.Enabled = True Then
        cmb_addendumby.SetFocus
        End If
        Exit Sub
    
    End If
    
    
   '********SECURITY***************************
   
    CURUSRDET_QR(0) = Trim(cmb_addendumby.Text)
    CURUSRDET_RS.MoreResults
    CURUSRDET_RS.Requery
    If CURUSRDET_RS.RowCount > 0 Then
        If LOGINTYP > CURUSRDET_RS("usrtype") Then
          ''   no action as he can commit
        ElseIf LOGINTYP < CURUSRDET_RS("usrtype") Then
            MsgBox "Sorry! You're not an Authorized person to write an Addendum !!", vbInformation, "Pathlogy"
            cmd_save.SetFocus
            Exit Sub
        End If
        If CURUSRDET_RS("ADDENDUM") <> "YES" Then
            MsgBox "Sorry! You are not authorised to Write Addendum of this Modality"
            cmd_save.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Sorry! The Security details of this user is not available check up in DISLNKUSR"
        cmd_commit.SetFocus
        Exit Sub
    End If
    If Len(Trim(cmb_addendumby.Text)) = 0 Then
        MsgBox "The Addendum by field is EMPTY.", vbCritical
        cmb_addendumby.SetFocus
        Exit Sub
    End If
''********SECURITY***************************
    SELDADDENDUM_QR(0) = req_chglen(Trim(cmb_reqno.Text))
    SELDADDENDUM_RS.MoreResults
    SELDADDENDUM_RS.Requery
    If SELDADDENDUM_RS.RowCount > 0 Then
       SELDADDENDUM_RS.MoveLast
        INSDADDENDUM_QR(0) = req_chglen(Trim(cmb_reqno.Text))
        INSDADDENDUM_QR(1) = SELDADDENDUM_RS("ADDEND_NO") + 1
        INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        INSDADDENDUM_QR(3) = time1
        INSDADDENDUM_QR(4) = cmb_addendumby.Text
        INSDADDENDUM_QR(5) = Null
        INSDADDENDUM_QR(5) = txt_addendum.Text
        INSDADDENDUM_QR.Execute
        
        UPDTADNDMTIME_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        UPDTADNDMTIME_QR(1) = time1
        UPDTADNDMTIME_QR(2) = req_chglen(Trim(cmb_reqno.Text))
        UPDTADNDMTIME_QR.Execute
    Else
        INSDADDENDUM_QR(0) = req_chglen(Trim(cmb_reqno.Text))
        INSDADDENDUM_QR(1) = "1"
        INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        INSDADDENDUM_QR(3) = time1
        INSDADDENDUM_QR(4) = cmb_addendumby.Text
        INSDADDENDUM_QR(5) = Null
        INSDADDENDUM_QR(5) = txt_addendum.Text
        INSDADDENDUM_QR.Execute
        
        UPDTADNDMTIME_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        UPDTADNDMTIME_QR(1) = time1
        UPDTADNDMTIME_QR(2) = req_chglen(Trim(cmb_reqno.Text))
        UPDTADNDMTIME_QR.Execute
        
    End If
    MsgBox "Addendum added Successfully!"
    
End If
rdb.CommitTrans
'MsgBox "Record is Saved"
btnenabl ("INI")

Exit Sub
MSG:
    rdb.RollbackTrans
'MsgBox Err.Description
End Sub
Public Function RNGFLG(RNG As String, VALE As String) As String
Dim RNGH As Double
Dim RNGL As Double
        RNGH = Mid(RNG, (InStr(1, RNG, "-") + 1))
        RNGL = Mid(RNG, 1, (InStr(1, RNG, "-") - 1))
        
        
If InStr(1, VALE, ">") Or InStr(1, VALE, "<") Then
VALE = Mid(VALE, 2)
End If

        If VALE < RNGL Then
            RNGFLG = "L"
        End If
        If VALE > RNGH Then
            RNGFLG = "H"
        End If
        If VALE <= RNGH And VALE >= RNGL Then
            RNGFLG = "N"
        End If
End Function



Public Sub ENABL(a As String)
    If a = "DIS" Then
        TXT_HB.Enabled = False
        TXT_RBC.Enabled = False
        TXT_TWBC.Enabled = False
        txt_platelets.Enabled = False
        txt_neutrophils.Enabled = False
        txt_lymphocytes.Enabled = False
        txt_monocytes.Enabled = False
        txt_eosinophils.Enabled = False
        txt_basophils.Enabled = False
        txt_ntphabs.Enabled = False
        txt_lymabs.Enabled = False
        txt_mncyteabs.Enabled = False
        txt_esnabs.Enabled = False
        txt_bsphabs.Enabled = False
        txt_immaturegranulocytes.Enabled = False
        TXT_HCT.Enabled = False
        TXT_MCH.Enabled = False
        TXT_MCV.Enabled = False
        TXT_MCV.Enabled = False
        TXT_MCHC.Enabled = False
        TXT_MPV.Enabled = False
        TXT_RDW.Enabled = False
        TXT_ABNORMAL.Enabled = False
        txt_mneutrophils.Enabled = False
        txt_mlymphocytes.Enabled = False
        txt_prolympho.Enabled = False 'Shraddha Pandit 16/10/08
        txt_mmonocytes.Enabled = False
        txt_meosinophils.Enabled = False
        txt_mbasophills.Enabled = False
        txt_metamyelocytes.Enabled = False
        txt_myelocytes.Enabled = False
        TXT_PROMYELOCYTESWBC.Enabled = False
        txt_promonocyte.Enabled = False
        txt_blasts.Enabled = False
        txt_plasmacells.Enabled = False
        txt_alymphocytes.Enabled = False
        'txt_prolymphocytes.Enabled = False
        txt_Sezarycells.Enabled = False
        txt_hairycells.Enabled = False
        txt_othercells.Enabled = False
        TXT_NORMOBLASTS.Enabled = False
        TXT_OTHERCELLSRBC.Enabled = False
        txt_rbcmorphology.Enabled = False
        txt_crc.Enabled = False 'Shraddha Pandit 20/11/08
        TXT_LP.Enabled = False  'Shraddha Pandit 17/10/08
        txt_rpi.Enabled = False 'Shraddha Pandit 20/11/08
        txt_empc.Enabled = False 'Shraddha Pandit 20/11/08
        txt_othertype.Enabled = False  'Shraddha Pandit 21/10/08
        txt_malarialparasite.Enabled = False  'Shraddha Pandit 21/10/08
        txt_cbc_comment.Enabled = False 'Shraddha Kesarkar 01/03/11
        txt_rc.Enabled = False
        Text1.Enabled = False
        txt_reticcount.Enabled = False
        txt_malarialparasite.Enabled = False
        txt_esr.Enabled = False '18/11/08
        'TXT_MORPHOLOGY.Enabled = False
        CMB_MORPHO.Enabled = False
        TXT_ADEQUACY.Enabled = False
        txt_inr.Enabled = False
        TXT_PTTK.Enabled = False
        TXT_FIBRINOGEN.Enabled = False
        TXT_OTHERS.Enabled = False
        TXT_PT(0).Enabled = False
        TXT_TT.Enabled = False
        TXT_FDPDMIER.Enabled = False
        txt_ptctl.Enabled = False
        txt_apttctl.Enabled = False
        txt_ttctl.Enabled = False
        txt_lapscore.Enabled = False
        txt_lapctl.Enabled = False
        TXT_MPO.Enabled = False
        cmb_mpo.Enabled = False
        cmbmpo.Enabled = False 'Added by Shraddha Pandit on 20/11/08
        cmb_nse.Enabled = False
        TXT_PAS.Enabled = False
        TXT_NSE.Enabled = False
        TXT_TRAP.Enabled = False
        cmb_enterby.Enabled = False
        cmb_finalby.Enabled = False
        txt_impression.Enabled = False
    End If
    If a = "ENA" Then
        TXT_HB.Enabled = True
        TXT_RBC.Enabled = True
        TXT_TWBC.Enabled = True
        txt_platelets.Enabled = True
        txt_neutrophils.Enabled = True
        txt_lymphocytes.Enabled = True
        txt_monocytes.Enabled = True
        txt_eosinophils.Enabled = True
        txt_basophils.Enabled = True
        txt_ntphabs.Enabled = True
        txt_lymabs.Enabled = True
        txt_mncyteabs.Enabled = True
        txt_esnabs.Enabled = True
        txt_bsphabs.Enabled = True
        txt_immaturegranulocytes.Enabled = True
        TXT_HCT.Enabled = True
        TXT_MCH.Enabled = True
        TXT_MCV.Enabled = True
        TXT_MCV.Enabled = True
        TXT_MCHC.Enabled = True
        TXT_MPV.Enabled = True
        TXT_RDW.Enabled = True
        TXT_ABNORMAL.Enabled = True
        txt_mneutrophils.Enabled = True
        txt_mlymphocytes.Enabled = True
        txt_prolympho.Enabled = True 'Shraddha Pandit 16/10/08
        txt_mmonocytes.Enabled = True
        txt_meosinophils.Enabled = True
        txt_mbasophills.Enabled = True
        txt_metamyelocytes.Enabled = True
        txt_myelocytes.Enabled = True
        TXT_PROMYELOCYTESWBC.Enabled = True
        txt_promonocyte.Enabled = True
        txt_blasts.Enabled = True
        txt_plasmacells.Enabled = True
        txt_alymphocytes.Enabled = True
        'txt_prolymphocytes.Enabled = True
        txt_Sezarycells.Enabled = True
        txt_hairycells.Enabled = True
        txt_othercells.Enabled = True
        TXT_NORMOBLASTS.Enabled = True
        TXT_OTHERCELLSRBC.Enabled = True
        txt_rbcmorphology.Enabled = True
        txt_crc.Enabled = True 'Shrddha Pandit 20/11/08
        cmbmpo.Enabled = True 'Added by Shraddha Pandit on 20/11/08
        TXT_LP.Enabled = True 'Shraddha Pandit 17/10/08
        txt_rpi.Enabled = True 'Shraddha Pandit 20/11/08
        txt_empc.Enabled = True 'Added by Shraddha Pandit on 20/11/08
        txt_othertype.Enabled = True 'Shraddha Pandit 21/10/08
        txt_malarialparasite.Enabled = True 'Shraddha Pandit 21/10/08
        txt_cbc_comment.Enabled = True 'Shraddha Kesarkar 01/03/11
        txt_rc.Enabled = True
        Text1.Enabled = True
        txt_reticcount.Enabled = True
        txt_malarialparasite.Enabled = True
        txt_esr.Enabled = True '18/11/08
        'TXT_MORPHOLOGY.Enabled = True
        CMB_MORPHO.Enabled = True
        TXT_ADEQUACY.Enabled = True
        txt_inr.Enabled = True
        TXT_PTTK.Enabled = True
        TXT_FIBRINOGEN.Enabled = True
        TXT_OTHERS.Enabled = True
        TXT_PT(0).Enabled = True
        TXT_TT.Enabled = True
        TXT_FDPDMIER.Enabled = True
        txt_ptctl.Enabled = True
        txt_apttctl.Enabled = True
        txt_ttctl.Enabled = True
        txt_lapscore.Enabled = True
        txt_lapctl.Enabled = True
        TXT_MPO.Enabled = True
        cmb_mpo.Enabled = True
        cmb_nse.Enabled = True
        TXT_PAS.Enabled = True
        TXT_NSE.Enabled = True
        TXT_TRAP.Enabled = True
        cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
        txt_impression.Enabled = True
    End If
    
End Sub

Private Sub Option3_Click()
'commented by shweta 04-02-2014 as told by kalpana mam(heamat dept. instruction)
'txt_ptctl.Text = "13.6"
'txt_apttctl.Text = "30.0"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'PREVTAB = PreviousTab
'MsgBox (SSTab1.Tab)
If SSTab1.Tab = 2 Then
'txt_ptctl.Text = 14
'txt_apttctl.Text = 29.5

End If
txt_1.Visible = False

PREVTAB = SSTab1.Tab
'************* Shraddha Pandit 30/03/09 ******************
If SSTab1.Caption = "Suface Marker" And ADDENDUMFLAG = True Then
    txt_addendum.Enabled = True
    cmb_addendumby.Enabled = True
    cmd_save.Enabled = True
End If
'*********************************************************
If SSTab1.Tab = 4 Then
     Me.Hide
     'If condition added by Shraddha on 04/12/09
     If txt_VIEW.Text = True Or txt_SAVE.Text = True Then
        SURFACEFORM.ShowProperOptions
        SURFACEFORM.opt_PrePBSC_CD3 = False
        SURFACEFORM.opt_PBSCharvest_CD3 = False
        SURFACEFORM.opt_BMharvest_CD3 = False
            
        SURFACEFORM.opt_Tcell_CD4 = False
        SURFACEFORM.opt_Tcell_CD8 = False
        
        SURFACEFORM.opt_PrePBSC_34 = False
        SURFACEFORM.opt_PBSCharvest_34 = False
        SURFACEFORM.opt_BMharvest_34 = False
     End If
     If txt_NEW.Text = True Then
        SURFACEFORM.opt_PrePBSC_CD3.Enabled = True
        SURFACEFORM.opt_PBSCharvest_CD3.Enabled = True
        SURFACEFORM.opt_BMharvest_CD3.Enabled = True
            
        SURFACEFORM.opt_Tcell_CD4.Enabled = True
        SURFACEFORM.opt_Tcell_CD8.Enabled = True
        
        SURFACEFORM.opt_PrePBSC_34.Enabled = True
        SURFACEFORM.opt_PBSCharvest_34.Enabled = True
        SURFACEFORM.opt_BMharvest_34.Enabled = True
        
        '*****
        SURFACEFORM.opt_PrePBSC_CD3 = False
        SURFACEFORM.opt_PBSCharvest_CD3 = False
        SURFACEFORM.opt_BMharvest_CD3 = False
            
        SURFACEFORM.opt_Tcell_CD4 = False
        SURFACEFORM.opt_Tcell_CD8 = False
        
        SURFACEFORM.opt_PrePBSC_34 = False
        SURFACEFORM.opt_PBSCharvest_34 = False
        SURFACEFORM.opt_BMharvest_34 = False
        '*****
     End If
     SURFACEFORM.show
Else
    txt_addendum.Enabled = False
    cmb_addendumby.Enabled = False
End If
 
End Sub

Private Sub Text1_GotFocus()
On Error GoTo m11:

w_pltmorphology.Visible = False
Set d_pltmorphology = w_pltmorphology.ActiveDocument
d_pltmorphology.Content = ""
d_pltmorphology.Content = Text1  'txt_pltmorphology
w_pltmorphology.Caption = "PLATELET MORPHOLOGY"
'w_frozen.WindowState = wdWindowStateMaximize
w_pltmorphology.WindowState = wdWindowStateNormal
d_pltmorphology.Content = Text1 'txt_pltmorphology
w_pltmorphology.Visible = True
w_pltmorphology.Activate
m11:
On Error Resume Next
Set w_pltmorphology = CreateObject("Word.Application")
w_pltmorphology.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
w_pltmorphology.Caption = "plt MORPHOLOGY"
'w_frozen.WindowState = wdWindowStateMaximize
w_pltmorphology.WindowState = wdWindowStateNormal
Set d_pltmorphology = w_pltmorphology.ActiveDocument
d_pltmorphology.Content = Trim(Text1)
w_pltmorphology.Visible = True
w_pltmorphology.Activate
End Sub

Private Sub Text1_LostFocus()
On Error GoTo m21:
Set d_pltmorphology = w_pltmorphology.ActiveDocument
Text1.Text = d_pltmorphology.Content
'flagimp = False
d_pltmorphology.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_pltmorphology.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_pltmorphology.Close (wdDoNotSaveChanges)
w_pltmorphology.quit
'W1.Visible = False
Exit Sub
m21:
MsgBox Err.Description
End Sub





Private Sub txt_1_GotFocus()
On Error GoTo m1
w_impression.Visible = False
Set d_impression = w_impression.ActiveDocument
d_impression.Content = ""
d_impression.Content = txt_1
w_impression.Caption = "COMMENTS"
'w_frozen.WindowState = wdWindowStateMaximize
w_impression.WindowState = wdWindowStateNormal
d_impression.Content = txt_1
w_impression.Visible = True
w_impression.Activate
m1:
On Error Resume Next
Set w_impression = CreateObject("Word.Application")
w_impression.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
w_impression.Caption = "COMMENTS"
'w_frozen.WindowState = wdWindowStateMaximize
w_impression.WindowState = wdWindowStateNormal
Set d_impression = w_impression.ActiveDocument
d_impression.Content = Trim(txt_1)
w_impression.Visible = True
w_impression.Activate
End Sub

Private Sub txt_1_lostfocus()
On Error GoTo m2:
Set d_impression = w_impression.ActiveDocument
txt_1.Text = d_impression.Content
'flagimp = False
d_impression.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_impression.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_impression.Close (wdDoNotSaveChanges)
w_impression.quit
'W1.Visible = False
Exit Sub
m2:

End Sub

Private Sub txt_addendum_GotFocus()
On Error GoTo m1:

'************************************
W_ADDENDUM.Visible = False
Set d_ADDENDUM = W_ADDENDUM.ActiveDocument
d_ADDENDUM.Content = ""
d_ADDENDUM.Content = txt_addendum
W_ADDENDUM.Caption = "ADDENDUM"
W_ADDENDUM.WindowState = wdWindowStateMaximize
d_ADDENDUM.Content = txt_addendum
W_ADDENDUM.Visible = True
W_ADDENDUM.Activate
m1:
On Error Resume Next
Set W_ADDENDUM = CreateObject("Word.Application")
W_ADDENDUM.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
W_ADDENDUM.Caption = "ADDENDUM"
W_ADDENDUM.WindowState = wdWindowStateMaximize
Set d_ADDENDUM = W_ADDENDUM.ActiveDocument
txt_addendum.Text = ""
d_ADDENDUM.Content = Trim(txt_addendum.Text)
If ADDENDUMFLAG = True Then
    d_ADDENDUM.Content = ""
End If
W_ADDENDUM.Visible = True
W_ADDENDUM.Activate
End Sub

Private Sub txt_addendum_LostFocus()
On Error GoTo m2:
Set d_ADDENDUM = W_ADDENDUM.ActiveDocument
txt_addendum.Text = d_ADDENDUM.Content
'flagimp = False
d_ADDENDUM.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_ADDENDUM.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_ADDENDUM.Close (wdDoNotSaveChanges)
W_ADDENDUM.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description

End Sub



Private Sub txt_basophils_LostFocus()
    txt_bsphabs.Text = (TXT_TWBC.Text) * ((txt_basophils.Text) / 100)
End Sub


Private Sub TXT_COMMENTS_GotFocus()
txt_1.Visible = False
If SSTab1.Tab = 4 Then
'On Error GoTo m1:
txt_1.Visible = True
txt_1.Enabled = True
txt_1.SetFocus
'w_impression.Visible = False
'Set d_impression = w_impression.ActiveDocument
'd_impression.Content = ""
'd_impression.Content = txt_comments
'w_impression.Caption = "COMMENTS"
''w_frozen.WindowState = wdWindowStateMaximize
'w_impression.WindowState = wdWindowStateNormal
'd_impression.Content = txt_comments
'w_impression.Visible = True
'w_impression.Activate
'm1:
'On Error Resume Next
'Set w_impression = CreateObject("Word.Application")
'w_impression.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
'w_impression.Caption = "COMMENTS"
''w_frozen.WindowState = wdWindowStateMaximize
'w_impression.WindowState = wdWindowStateNormal
'Set d_impression = w_impression.ActiveDocument
'd_impression.Content = Trim(txt_comments)
'w_impression.Visible = True
'w_impression.Activate
Exit Sub
End If

If Trim(txt_comments.Text) = "" Then
    txt_comments.clear
    txt_comments.AddItem "# Kindly note that the results might be erroneous as sample collected 4.5 ml is less than the reuired volume 4.5 ml. "
    txt_comments.AddItem ""
    txt_comments.AddItem "# Sample Haemolysed"
    txt_comments.AddItem "# Sample Icteric"
    txt_comments.AddItem "# Sample Lipemic"
    txt_comments.AddItem "# Reanalysis Done"
    txt_comments.AddItem "# Please Repeat Collection"
    txt_comments.AddItem "# Sample Run On Extended Mode"
    txt_comments.AddItem "Manual Platelet Count matches with count given by analyzer."
    '***** Added by Shraddha on 30/11/2011 *****
    txt_comments.AddItem "Platelets seen in clumps, advice repeat collection in sodium citrate bulb for correct platelet count"
    txt_comments.AddItem "Total count very low DC of 50 cells"
    '********************************************
    'txt_comments.AddItem "# Serum lipaemic"
    'txt_comments.AddItem "# Serum highly lipaemic"
    'txt_comments.AddItem "# Reading out of range. Pl repeat the sample"
   ' txt_comments.AddItem "# Electrolytes reported from E-Lab"
   ' txt_comments.AddItem "# Please send Heparinised Sample to E-Lab"
   ' txt_comments.AddItem "# Out of Range (High/Low)sample ?,Please repeat the sample"
    'txt_comments.AddItem "# Report Attached."
    'txt_comments.AddItem ""
    
Else
    txt_comments.clear
    
    txt_comments.AddItem " " & Trim(SELDBMAREPORT_RS("COMMENTS").Value)
    txt_comments.AddItem ""
    txt_comments.AddItem "# Kindly note that the results might be erroneous as sample collected 4.5 ml is less than the reuired volume 4.5 ml. "
    txt_comments.AddItem "# Sample Haemolysed"
    txt_comments.AddItem "# Sample Icteric"
    txt_comments.AddItem "# Sample Lipemic"
    txt_comments.AddItem "# Reanalysis Done"
    txt_comments.AddItem "# Please Repeat Collection"
    txt_comments.AddItem "# Sample Run On Extended Mode"
    txt_comments.AddItem "Manual Platelet ount matches with count given by analyzer."
    'txt_comments.AddItem ""
    
End If


End Sub
'Private Sub TXT_COMMENTS_GotFocus()
'On Error GoTo m1:
'
'w_comments.Visible = False
'Set d_comments = w_comments.ActiveDocument
'd_comments.Content = ""
'd_comments.Content = txt_comments
'w_comments.Caption = "Comments"
''w_frozen.WindowState = wdWindowStateMaximize
'w_comments.WindowState = wdWindowStateNormal
'd_comments.Content = txt_comments
'w_comments.Visible = True
'w_comments.Activate
'm1:
'On Error Resume Next
'Set w_comments = CreateObject("Word.Application")
'w_comments.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
'w_comments.Caption = "Comments"
''w_frozen.WindowState = wdWindowStateMaximize
'w_comments.WindowState = wdWindowStateNormal
'Set d_comments = w_comments.ActiveDocument
'd_comments.Content = Trim(txt_comments)
'w_comments.Visible = True
'w_comments.Activate
'End Sub

Private Sub txt_comments_lostfocus()
If txt_1.Visible = True Then
Exit Sub
Else

On Error GoTo m2:
Set d_impression = w_impression.ActiveDocument
txt_comments.Text = d_impression.Content
'flagimp = False
d_impression.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_impression.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_impression.Close (wdDoNotSaveChanges)
w_impression.quit
'W1.Visible = False
Exit Sub
End If
m2:
MsgBox Err.Description
End Sub

'Private Sub txt_comments_lostfocus()
'On Error GoTo m4:
''Added by Milind on 02/02/06 for removing Comments from Report when it is blank
'txt_comments.Text = d_comments.Content
'Set d_comments = w_comments.ActiveDocument
'
'If blank = True Then
'    l = Len(d_comments.Content)
'    If l <= 4 Then
'        txt_comments.Text = ""
'    Else
'        txt_comments.Text = d_comments.Content
'     End If
'End If
'
''upto here
'
'd_comments.Content = ""
'd_comments.Close (wdDoNotSaveChanges)
'w_comments.quit
'Exit Sub
'm4:
'MsgBox Err.Description
'
'End Sub

Private Sub txt_eosinophils_LostFocus()
    txt_esnabs.Text = (TXT_TWBC.Text) * ((txt_eosinophils.Text) / 100)
End Sub





Private Sub TXT_FDPDMIER_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
KeyAscii = 0
End If
End Sub

Private Sub TXT_FDPDMIER_LostFocus()
'If Trim(TXT_FDPDMIER.Text) <> "" Then
'    If InStr(1, TXT_FDPDMIER.Text, ">") = 0 Then
'        If InStr(1, TXT_FDPDMIER.Text, "<") = 0 Then
'            MsgBox "Enter the value OF D-DIMER in format >  <"
'            TXT_FDPDMIER.SetFocus
'            Exit Sub
'        End If
'    End If
'End If
'Exit Sub
End Sub

Private Sub TXT_FIBRINOGEN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'SendKeys vbTab
KeyAscii = 0
End If
End Sub

Private Sub txt_impression_gotfocus()
On Error GoTo m1:

w_impression.Visible = False
Set d_impression = w_impression.ActiveDocument
d_impression.Content = ""
d_impression.Content = txt_impression
w_impression.Caption = "Impression"
'w_frozen.WindowState = wdWindowStateMaximize
w_impression.WindowState = wdWindowStateNormal
d_impression.Content = txt_impression
w_impression.Visible = True
w_impression.Activate
m1:
On Error Resume Next
Set w_impression = CreateObject("Word.Application")
w_impression.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
w_impression.Caption = "Impression"
'w_frozen.WindowState = wdWindowStateMaximize
w_impression.WindowState = wdWindowStateNormal
Set d_impression = w_impression.ActiveDocument
d_impression.Content = Trim(txt_impression)
w_impression.Visible = True
w_impression.Activate
End Sub

Private Sub txt_impression_lostfocus()
On Error GoTo m2:
Set d_impression = w_impression.ActiveDocument
txt_impression.Text = d_impression.Content
'flagimp = False
d_impression.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_impression.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_impression.Close (wdDoNotSaveChanges)
w_impression.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description
End Sub



Private Sub cmd_back_Click()
'Me.Hide
'DIS_MAIN.show

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

Public Sub TCLEAR()

        TXT_HB.Text = ""
        TXT_RBC.Text = ""
        TXT_TWBC.Text = ""
        txt_platelets.Text = ""
        txt_neutrophils.Text = ""
        txt_lymphocytes.Text = ""
        txt_monocytes.Text = ""
        txt_eosinophils.Text = ""
        txt_basophils.Text = ""
        txt_ntphabs.Text = ""
        txt_lymabs.Text = ""
        txt_mncyteabs.Text = ""
        txt_esnabs.Text = ""
        txt_bsphabs.Text = ""
        txt_immaturegranulocytes.Text = ""
        TXT_HCT.Text = ""
        TXT_MCH.Text = ""
        TXT_MCV.Text = ""
        TXT_MCV.Text = ""
        TXT_MCHC.Text = ""
        TXT_MPV.Text = ""
        TXT_RDW.Text = ""
        TXT_ABNORMAL.Text = ""
        txt_mneutrophils.Text = ""
        txt_mlymphocytes.Text = ""
        txt_prolympho.Text = "" 'Shraddha Pandit 16/10/08
        txt_mmonocytes.Text = ""
        txt_meosinophils.Text = ""
        txt_mbasophills.Text = ""
        txt_metamyelocytes.Text = ""
        txt_myelocytes.Text = ""
        TXT_PROMYELOCYTESWBC.Text = ""
        txt_promonocyte.Text = ""
        txt_blasts.Text = ""
        cmbmpo.Text = ""
        txt_plasmacells.Text = ""
        txt_alymphocytes.Text = ""
        txt_prolymphocytes.Text = ""
        txt_Sezarycells.Text = ""
        txt_hairycells.Text = ""
        txt_othercells.Text = ""
        TXT_NORMOBLASTS.Text = ""
        TXT_OTHERCELLSRBC.Text = ""
        txt_rbcmorphology.Text = ""
        txt_empc.Text = "" 'Added by Shraddha Pandit on 20/11/08
        TXT_LP.Text = "" 'Shraddha Pandit 17/10/08
        txt_othertype = "" 'Shraddha Pandit 21/10/08
        txt_crc.Text = "" 'Shraddha Pandit 20/11/08
        txt_rc.Text = ""
        Text1.Text = " "
        txt_reticcount.Text = ""
        txt_malarialparasite.Text = ""
        txt_esr.Text = "" '18/11/08
        'TXT_MORPHOLOGY.Text = ""
        TXT_ADEQUACY.Text = ""
        txt_inr.Text = ""
        TXT_PTTK.Text = ""
        TXT_FIBRINOGEN.Text = ""
        TXT_OTHERS.Text = ""
        TXT_PT(0).Text = ""
        TXT_TT.Text = ""
        TXT_FDPDMIER.Text = ""
        txt_ptctl.Text = ""
        txt_apttctl.Text = ""
        txt_ttctl.Text = ""
        txt_lapscore.Text = ""
        txt_lapctl.Text = ""
        TXT_MPO.Text = ""
        cmb_mpo.Text = ""
        cmbmpo.Text = "" 'Added by Shraddha Pandit on 20/11/08
        cmb_nse.Text = ""
        TXT_PAS.Text = ""
        TXT_NSE.Text = ""
        TXT_TRAP.Text = ""
        'cmb_enterby.Text = ""
        'cmb_finalby.Text = ""
        regcbo.Text = ""
        
        txt_impression.Text = ""
        txt_addendum.Text = ""
        '' SURFACE FORM TEXDT CLEAR
        SURFACEFORM.TXT_ANTIMPO = " "
        SURFACEFORM.TXT_CD10 = " "
        SURFACEFORM.txt_cd103 = " "
        SURFACEFORM.TXT_CD117 = " "
        SURFACEFORM.TXT_CD117 = " "
        SURFACEFORM.TXT_CD11C = " "
        SURFACEFORM.TXT_CD11C = " "
        SURFACEFORM.txt_cd13 = " "
        SURFACEFORM.TXT_CD138 = " "
        SURFACEFORM.txt_cd14 = " "
        SURFACEFORM.txt_CD15 = " "
        SURFACEFORM.TXT_CD19 = " "
        SURFACEFORM.TXT_CD2 = " "
        SURFACEFORM.TXT_CD20 = ""
        SURFACEFORM.TXT_CD22 = ""
        SURFACEFORM.TXT_CD23 = ""
        SURFACEFORM.TXT_CD25 = ""
        SURFACEFORM.TXT_CD3 = " "
        SURFACEFORM.txt_CD33 = ""
        SURFACEFORM.TXT_CD34 = " "
        SURFACEFORM.txt_cd36 = " "
        SURFACEFORM.TXT_CD38 = " "
        SURFACEFORM.TXT_CD4 = " "
        SURFACEFORM.txt_cd41 = " "
        SURFACEFORM.TXT_CD42 = " "
        SURFACEFORM.TXT_CD5 = " "
        SURFACEFORM.TXT_CD56 = " "
        SURFACEFORM.TXT_CD61 = " "
        SURFACEFORM.TXT_CD68 = " "
        SURFACEFORM.TXT_CD7 = " "
        SURFACEFORM.txt_CD79B = " "
        SURFACEFORM.txt_CD79B = " "
        SURFACEFORM.TXT_CD8 = ""
        SURFACEFORM.txt_cytocd221(1) = " "
        SURFACEFORM.TXT_CYTOCD3 = " "
        SURFACEFORM.txt_cytocd79a = " "
        SURFACEFORM.TXT_CYTOMU = " "
        SURFACEFORM.TXT_DBA44 = " "
        SURFACEFORM.TXT_FMC7 = " "
        SURFACEFORM.txt_Glycophorina = " "
        SURFACEFORM.TXT_HLADR = " "
        SURFACEFORM.txt_igm = " "
        SURFACEFORM.txt_kappa = " "
        SURFACEFORM.txt_lanbda = " "
        SURFACEFORM.txt_smig = " "
        SURFACEFORM.TXT_TDT = " "
        SURFACEFORM.TXT_TDT = " "
        
End Sub

Private Sub SPREPORT()
Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim I As Integer
On Error GoTo prtmsg:
DoEvents
Dim VAR_REQNNO As String
VAR_REQNNO = req_chglen(Trim(REQNO1))
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'seldrequisitn_qr10.QueryTimeout = 150
If seldrequisitn_flag11 = False Then
    seldrequisitn_qr11(0) = MaskEdBox1.FormattedText
    Set seldrequisitn_rs11 = seldrequisitn_qr11.OpenResultset(1, 3)
    seldrequisitn_flag11 = True
Else
    seldrequisitn_qr11(0) = MaskEdBox1.FormattedText
    seldrequisitn_rs11.MoreResults
    seldrequisitn_rs11.Requery
End If
If seldrequisitn_rs11.RowCount > 0 Then
    prpathno = ""
    'CMB_PATHNO.clear
    While Not seldrequisitn_rs11.EOF
        prpathno = prpathno & " " & seldrequisitn_rs11("pathno")
        'CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & "$" & seldrequisitn_rs10("REQNO")
        seldrequisitn_rs11.MoveNext
    Wend
    'txt_prepathreqno.Text = prpathno
End If
'If SELDADDENDUM_FLAG = False Then
 '   SELDADDENDUM_QR(0) = VAR_REQNNO
    'Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
    'SELDADDENDUM_FLAG = True
'Else
    SELDADDENDUM_QR(0) = VAR_REQNNO
    SELDADDENDUM_RS.MoreResults
    SELDADDENDUM_RS.Requery
'End If

For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I


str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
CrystalReport1.Formulas(0) = "REQNO = '" & Trim(REQNO1) & "'"
If Opt_patient.Value = True Then
        If SELRREQNORS("COMMITFLAG") = "Y" Then
            If SELDADDENDUM_RS.RowCount > 0 Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
            End If
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
        End If
        SELPATHDREQUISITN_QR(0) = VAR_REQNNO
        Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
        If selpathdrequisitn_rs.RowCount > 0 Then
             CrystalReport1.Formulas(1) = "PREPATH = 'Prev. Path No." & Trim(prpathno) & "'"
        Else
             CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        selta1012_qr(0) = MaskEdBox1.FormattedText
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        If selta1012_rs.RowCount > 0 Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        Else
            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
        End If
        
        Dim z
        Dim Y
        Dim w
        Dim u
        
        SELMA301QR(0) = selpathdrequisitn_rs("SITE")
        Set SELMA301RS = SELMA301QR.OpenResultset(1, 3)
        If SELMA301RS.RowCount > 0 Then
            CrystalReport1.Formulas(3) = "SIT = '" & Trim(SELMA301RS("SITEDESC")) & "'"
        Else
            CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            If Trim(selpathdrequisitn_rs("OPPROCEDUR")) <> "" Then
                z = InStr(1, Trim(selpathdrequisitn_rs("OPPROCEDUR")), "'")
                If z <> 0 Then
                    Y = Trim(selpathdrequisitn_rs("OPPROCEDUR"))
                    u = Len(Y)
                    X = Left(Y, z - 1)
                    w = Right(Y, u - (z))
                    Y = X & w
                    CrystalReport1.Formulas(4) = "OPPRO = '" & Y & "'"
                Else
                    CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(selpathdrequisitn_rs("OPPROCEDUR")) & "'"
                End If
            Else
                CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(selpathdrequisitn_rs("OPPROCEDUR")) & "'"
            End If
        Else
            CrystalReport1.Formulas(4) = "OPPRO = '" & "" & "'"
        End If
        str1 = ""
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        SELDFREPORT_QR(0) = VAR_REQNNO
        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
        If SELDFREPORT_RS.RowCount > 0 Then
            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
                CrystalReport1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
            Else
                CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
                CrystalReport1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
            Else
                CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
            Else
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
            Else
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
                CrystalReport1.Formulas(12) = "IMPN = '" & "Impression" & "'"
            Else
                CrystalReport1.Formulas(12) = "IMPN = '" & "" & "'"
            End If
        End If
        
        CrystalReport1.Formulas(8) = "USRID = '" & str & "'"
        
        'CrystalReport1.Formulas(10) = "GROSDESC = '" & "Gross description" & "'"
        'CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
        CrystalReport1.WindowState = crptMaximized
      '  CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    Else
        SELEMPBASINFO_QR(0) = MaskEdBox1.Text
        SELEMPBASINFO_RS.MoreResults
        SELEMPBASINFO_RS.Requery
'        If SELEMPBASINFO_RS.RowCount > 0 Then
'            txt_sex.Text = SELEMPBASINFO_RS("GENDER")
'            txt_name.Text = Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME"))
'            txt_age = DateDiff("YYYY", SELEMPBASINFO_RS("DOB"), CStr(DATE11))
'        End If

        If SELRREQNORS("COMMITFLAG") = "Y" Then
            If SELDADDENDUM_RS.RowCount > 0 Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADDSTAFF.RPT"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
            End If
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
        End If
        SELPATHDREQUISITN_QR(0) = VAR_REQNNO
        Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
        If selpathdrequisitn_rs.RowCount > 0 Then
             CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(selpathdrequisitn_rs("prepathno")) & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        selta1012_qr(0) = MaskEdBox1.FormattedText
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        If selta1012_rs.RowCount > 0 Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        Else
            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
        End If
        
        
        SELMA301QR(0) = selpathdrequisitn_rs("SITE")
        Set SELMA301RS = SELMA301_QR.OpenResultset(1, 3)
        If SELMA301RS.RowCount > 0 Then
            CrystalReport1.Formulas(3) = "SIT = '" & Trim(SELMA301RS("SITEDESC")) & "'"
        Else
            CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
        End If
        If selpathdrequisitn_rs.RowCount > 0 Then
            CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(selpathdrequisitn_rs("OPPROCEDUR")) & "'"
        Else
            CrystalReport1.Formulas(4) = "OPPRO = '" & "" & "'"
        End If
        str1 = ""
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        SELDFREPORT_QR(0) = VAR_REQNNO
        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
        If SELDFREPORT_RS.RowCount > 0 Then
            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
                CrystalReport1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
            Else
                CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
                CrystalReport1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
            Else
                CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
            Else
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
            Else
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
                CrystalReport1.Formulas(12) = "IMPN = '" & "Impression" & "'"
            Else
                CrystalReport1.Formulas(12) = "IMPN = '" & "" & "'"
            End If
         
        End If
        CrystalReport1.Formulas(8) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(13) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(14) = "DOB = '" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT")) & "'"
        CrystalReport1.Formulas(15) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    End If
'CrystalReport1.WindowState = crptMaximized
'CrystalReport1.Action = 1
Exit Sub
prtmsg: MsgBox ERROR(Err)
'txt_msg.Text = ""
'btnenabl ("print")
Exit Sub
MSG:
MsgBox ERROR(Err)
End Sub
Private Sub FAREPORT()
Dim VAR_REQNNO As String
VAR_REQNNO = req_chglen(Trim(REQNO1))
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'seldrequisitn_qr10.QueryTimeout = 300
'If seldrequisitn_flag10 = False Then
'    seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
'    Set seldrequisitn_rs10 = seldrequisitn_qr10.OpenResultset(1, 3)
'    seldrequisitn_flag10 = True
'Else
'    seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
'    seldrequisitn_rs10.MoreResults
'    seldrequisitn_rs10.Requery
'End If
'If seldrequisitn_rs10.RowCount > 0 Then
'    prpathno = ""
'    CMB_PATHNO.clear
'    While Not seldrequisitn_rs10.EOF
'        prpathno = prpathno & " " & seldrequisitn_rs10("pathno")
'        CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & "$" & seldrequisitn_rs10("REQNO")
'        seldrequisitn_rs10.MoveNext
'    Wend
'    txt_prepathreqno.Text = prpathno
'End If
'If SELDADDENDUM_FLAG = False Then
'    SELDADDENDUM_QR(0) = VAR_REQNNO
'    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'    SELDADDENDUM_FLAG = True
'Else
'    SELDADDENDUM_QR(0) = VAR_REQNNO
'    SELDADDENDUM_RS.MoreResults
'    SELDADDENDUM_RS.Requery
'End If

For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(REQNO1) & "'"

If Opt_patient.Value = True Then
        If SELRREQNORS("REQNSTAT") <> "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFA.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(1) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & SELRREQNORS("ENTERBY")
                CrystalReport1.Formulas(2) = "ENT = '" & str1 & "'"
             Else
    
                CrystalReport1.Formulas(1) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                CrystalReport1.Formulas(2) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADD.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
        End If
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    Else
        SELEMPBASINFO_QR(0) = MaskEdBox1.Text
        SELEMPBASINFO_RS.MoreResults
        SELEMPBASINFO_RS.Requery
        If SELRREQNORS("REQNSTAT") <> "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFASTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & SELRREQNORS("ENTERBY")
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADDSTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        End If
        CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT"))
        CrystalReport1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    End If
End Sub
Private Sub CGCNREPORT()
Dim VAR_REQNNO As String
VAR_REQNNO = req_chglen(Trim(REQNO1))
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'seldrequisitn_qr10.QueryTimeout = 300
'If seldrequisitn_flag10 = False Then
'    seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
'    Set seldrequisitn_rs10 = seldrequisitn_qr10.OpenResultset(1, 3)
'    seldrequisitn_flag10 = True
'Else
'    seldrequisitn_qr10(0) = MaskEdBox1.FormattedText
'    seldrequisitn_rs10.MoreResults
'    seldrequisitn_rs10.Requery
'End If
'SELPATHDREQUISITN_QR(0) = VAR_REQNNO
'Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
'
'If seldrequisitn_rs10.RowCount > 0 Then
'    prpathno = ""
'    CMB_PATHNO.clear
'    While Not seldrequisitn_rs10.EOF
'        prpathno = prpathno & " " & seldrequisitn_rs10("pathno")
'        CMB_PATHNO.AddItem seldrequisitn_rs10("PATHNO") & "$" & seldrequisitn_rs10("REQNO")
'        seldrequisitn_rs10.MoveNext
'    Wend
'    txt_prepathreqno.Text = prpathno
'End If
'If SELDADDENDUM_FLAG = False Then
'    SELDADDENDUM_QR(0) = VAR_REQNNO
'    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'    SELDADDENDUM_FLAG = True
'Else
'    SELDADDENDUM_QR(0) = VAR_REQNNO
'    SELDADDENDUM_RS.MoreResults
'    SELDADDENDUM_RS.Requery
'End If

'SELDADDENDUM_QR(0) = VAR_REQNNO
'SELDADDENDUM_RS.MoreResults
'SELDADDENDUM_RS.Requery

For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I
SELPATHDREQUISITN_QR(0) = VAR_REQNNO
Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"

If Opt_patient.Value = True Then
    If SELRREQNORS("REQNSTAT") <> "F" Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCG.RPT"
        'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(4) = "NOSLIDES='" & Trim(selpathdrequisitn_rs("LABNO")) & "'"
        If Trim(SELRREQNORS("PREPORTBY")) <> "" And Trim(SELRREQNORS("FREPORTBY")) <> "" And UCase(Trim(SELRREQNORS("PREPORTBY"))) <> UCase(Trim(SELRREQNORS("FREPORTBY"))) Then
            CrystalReport1.Formulas(11) = "P= '" & Trim(SELRREQNORS("PREPORTBY")) & "'"
            CrystalReport1.Formulas(12) = "F= '" & Trim(SELRREQNORS("FREPORTBY")) & "'"
            CrystalReport1.Formulas(13) = "PCAP= '" & "Cytotechnologist" & " '"
            CrystalReport1.Formulas(14) = "FCAP= '" & "Pathologist" & " '"
        End If
        If Trim(SELRREQNORS("PREPORTBY")) <> "" And Trim(SELRREQNORS("FREPORTBY")) = "" Then
            CrystalReport1.Formulas(11) = "PONLY= '" & Trim(SELRREQNORS("PREPORTBY")) & "'"
            'CrystalReport1.Formulas(12) = "F= '" & "" & "'"
            CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
        End If
        If Trim(SELRREQNORS("PREPORTBY")) <> "" And Trim(SELRREQNORS("FREPORTBY")) <> "" And UCase(Trim(SELRREQNORS("PREPORTBY"))) = UCase(Trim(SELRREQNORS("FREPORTBY"))) Then
            CrystalReport1.Formulas(11) = "PONLY= '" & Trim(SELRREQNORS("PREPORTBY")) & "'"
            CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
        End If
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADD.RPT"
        'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(12) = "NOSLIDES = " & "" & selpathdrequisitn_rs("LABNO")
    End If
Else
    If SELRREQNORS("REQNSTAT") <> "F" Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGSTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADDSTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
    End If
    CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
    CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
    CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREQNORS("REQNDT"))
End If
CrystalReport1.WindowState = crptMaximized
'CRYSTALREPORT1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CrystalReport1.Action = 1
End Sub

Private Sub BONEMARROW()
Dim VAR_REQNNO As String
VAR_REQNNO = req_chglen(Trim(REQNO1))
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
'''********BONE MARROW*************
For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I
    If SELDBMAREP_FLAG = False Then
        SELDBMAREP_QR(0) = VAR_REQNNO
        Set SELDBMAREP_RS = SELDBMAREP_QR.OpenResultset(1, 3)
        SELDBMAREP_FLAG = True
    Else
        SELDBMAREP_QR(0) = VAR_REQNNO
        SELDBMAREP_RS.MoreResults
        SELDBMAREP_RS.Requery
    End If
    If Opt_patient.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBM1.RPT"
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Not selta1012_rs.EOF Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & selta1012_rs("PRVDIAGDES") & "'"
        End If
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & SELRREQNORS("ENTERBY")
            If Len(Trim(SELRREQNORS("ENTERBY"))) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
            End If
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        If IsNull(SELDBMAREP_RS("CD2")) And IsNull(SELDBMAREP_RS("CD3")) And IsNull(SELDBMAREP_RS("CD4")) And IsNull(SELDBMAREP_RS("CD5")) And IsNull(SELDBMAREP_RS("CD7")) And IsNull(SELDBMAREP_RS("CD8")) And IsNull(SELDBMAREP_RS("CYTOCD3")) Then
            CrystalReport1.Formulas(13) = "tmarkers = '" & "" & "'"
        Else
            CrystalReport1.Formulas(13) = "tmarkers = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREP_RS("CD10")) And IsNull(SELDBMAREP_RS("CD19")) And IsNull(SELDBMAREP_RS("FMC7")) And IsNull(SELDBMAREP_RS("CD22")) And IsNull(SELDBMAREP_RS("CD23")) And IsNull(SELDBMAREP_RS("smig")) And IsNull(SELDBMAREP_RS("CD79B")) And IsNull(SELDBMAREP_RS("CD25")) And IsNull(SELDBMAREP_RS("CD11C")) And IsNull(SELDBMAREP_RS("cd103")) And IsNull(SELDBMAREP_RS("kappa")) And IsNull(SELDBMAREP_RS("lamda")) And IsNull(SELDBMAREP_RS("DBA44")) And IsNull(SELDBMAREP_RS("cytoEP22")) And IsNull(SELDBMAREP_RS("cytocd79a")) And IsNull(SELDBMAREP_RS("CYTOMU")) Then
            CrystalReport1.Formulas(14) = "bmarkers = '" & "" & "'"
        Else
            CrystalReport1.Formulas(14) = "bmarkers = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREP_RS("CD34")) And IsNull(SELDBMAREP_RS("HLADR")) And IsNull(SELDBMAREP_RS("CD38")) And IsNull(SELDBMAREP_RS("TDT")) And IsNull(SELDBMAREP_RS("CD138")) And IsNull(SELDBMAREP_RS("ANTIMPO")) Then
            CrystalReport1.Formulas(15) = "others = '" & "" & "'"
        Else
                CrystalReport1.Formulas(15) = "others = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREP_RS("cd13")) And IsNull(SELDBMAREP_RS("Cd14")) And IsNull(SELDBMAREP_RS("CD15")) And IsNull(SELDBMAREP_RS("CD33")) And IsNull(SELDBMAREP_RS("cd36")) And IsNull(SELDBMAREP_RS("cd4142")) And IsNull(SELDBMAREP_RS("Glycophoin")) And IsNull(SELDBMAREP_RS("CD68")) And IsNull(SELDBMAREP_RS("CD42")) And IsNull(SELDBMAREP_RS("CD117")) Then
           CrystalReport1.Formulas(16) = "mm = '" & "" & "'"
        Else
           CrystalReport1.Formulas(16) = "mm = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREP_RS("IRONSTAIN")) And IsNull(SELDBMAREP_RS("GIDEROBLAS")) And IsNull(SELDBMAREP_RS("sudanblack")) And IsNull(SELDBMAREP_RS("TOLUIDINEB")) Then
           CrystalReport1.Formulas(17) = "cci = '" & "" & "'"
        Else
           CrystalReport1.Formulas(17) = "cci = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREP_RS("myelopero")) And IsNull(SELDBMAREP_RS("nonspester")) And IsNull(SELDBMAREP_RS("trap")) Then
           CrystalReport1.Formulas(18) = "ccm = '" & "" & "'"
        Else
           CrystalReport1.Formulas(18) = "ccm = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREP_RS("IMPRINT")) And IsNull(SELDBMAREP_RS("abcellblas")) Then
           CrystalReport1.Formulas(19) = "im = '" & "" & "'"
        Else
           CrystalReport1.Formulas(19) = "im = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREP_RS("mcellulari")) And IsNull(SELDBMAREP_RS("morphology")) Then
           CrystalReport1.Formulas(20) = "ms = '" & "" & "'"
        Else
           CrystalReport1.Formulas(20) = "ms = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREP_RS("ecellulari")) And IsNull(SELDBMAREP_RS("maturation")) Then
           CrystalReport1.Formulas(21) = "es = '" & "" & "'"
        Else
           CrystalReport1.Formulas(21) = "es = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREP_RS("mature")) And IsNull(SELDBMAREP_RS("immature")) And IsNull(SELDBMAREP_RS("atypbinu")) And IsNull(SELDBMAREP_RS("plasmablas")) Then
           CrystalReport1.Formulas(22) = "pc = '" & "" & "'"
        Else
           CrystalReport1.Formulas(22) = "pc = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREP_RS("MLYMPHOCYT")) And IsNull(SELDBMAREP_RS("ALYMPHOCYT")) And IsNull(SELDBMAREP_RS("Plasmacell")) Then
           CrystalReport1.Formulas(23) = "ls = '" & "" & "'"
        Else
           CrystalReport1.Formulas(23) = "ls = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREP_RS("PROMYELOCY")) And IsNull(SELDBMAREP_RS("PROMYELOCY")) And IsNull(SELDBMAREP_RS("MMYELOCYTE")) And IsNull(SELDBMAREP_RS("polymorphs")) And IsNull(SELDBMAREP_RS("basophils")) And IsNull(SELDBMAREP_RS("eosinophil")) And IsNull(SELDBMAREP_RS("monocytes")) And IsNull(SELDBMAREP_RS("PROMONOCYT")) Then
           CrystalReport1.Formulas(24) = "mys = '" & "" & "'"
        Else
           CrystalReport1.Formulas(24) = "mys = '" & "abc" & "'"
        End If
        If IsNull(SELDBMAREP_RS("ECELLULARI")) And IsNull(SELDBMAREP_RS("MERATIO")) Then
           CrystalReport1.Formulas(25) = "mor = '" & "" & "'"
        Else
           CrystalReport1.Formulas(25) = "mor = '" & "abc" & "'"
        End If
    End If
    CrystalReport1.WindowState = crptMaximized
    'CRYSTALREPORT1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport1.Action = 1

End Sub
Private Sub BFREPORT()
For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I
selta1012_qr(0) = MaskEdBox1.FormattedText
Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)

Dim VAR_REQNNO As String
VAR_REQNNO = req_chglen(Trim(REQNO1))
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If
If Opt_patient.Value = True Then
             
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"

    
    
     If Trim(SELRREQNORS("MODALITYCD")) = "BF" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("ENTERBY"))
            If Len(Trim(SELRREQNORS("ENTERBY"))) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                 CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
            End If
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        
        
     End If
       
End If
If opt_staff.Value = True Then
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"

 
     If Trim(SELRREQNORS("MODALITYCD")) = "BF" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBFS.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & VAR_REQNNO & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        
        
     End If

End If
'***** STAFF
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CrystalReport1.Action = 1
End Sub
Private Sub HLREPORT()
Dim VAR_REQNNO As String
VAR_REQNNO = req_chglen(Trim(REQNO1))
If SELRREQNOFLAG = False Then
    SELRREQNOQR(0) = VAR_REQNNO
    Set SELRREQNORS = SELRREQNOQR.OpenResultset(1, 3)
    SELRREQNOFLAG = True
Else
    SELRREQNOQR(0) = VAR_REQNNO
    SELRREQNORS.MoreResults
    SELRREQNORS.Requery
End If

If SELDCBCREPORT_FLAG = False Then
    SELDCBCREPORT_QR(0) = VAR_REQNNO
    Set SELDCBCREPORT_RS = SELDCBCREPORT_QR.OpenResultset(1, 3)
    SELDCBCREPORT_FLAG = True
Else
    SELDCBCREPORT_QR(0) = VAR_REQNNO
    SELDCBCREPORT_RS.MoreResults
    SELDCBCREPORT_RS.Requery
End If
For I = 0 To 25
    CrystalReport1.Formulas(I) = ""
Next I
selta1012_qr(0) = SELRREQNORS("CASENO")
Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)

If Opt_patient.Value = True Then
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
    'If Trim(txt_modcode.Text) = "HL" Then
    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShl.RPT"
    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'       If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'       Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'       End If
    CrystalReport1.Formulas(2) = "PROVDIA = '" & selta1012_rs("PRVDIAGDES") & "'"
    If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
        CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
        str1 = ""
        str1 = "Entered By " & Trim(SELRREQNORS("enterby"))
        If Len(Trim(SELRREQNORS("enterby"))) > 0 Then
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
        End If
    Else
        CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
    End If
    CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
             If Trim(SELDCBCREPORT_RS("lap")) = "" And Trim(SELDCBCREPORT_RS("MPO")) = "" And Trim(SELDCBCREPORT_RS("SBB")) = "" And Trim(SELDCBCREPORT_RS("PSA")) = "" And Trim(SELDCBCREPORT_RS("NSE")) = "" And Trim(SELDCBCREPORT_RS("TRAD")) = "" Then
                'If CYTOCHEMFLAG = True Then
                CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
            Else
                CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
                'End If
            End If
            If Trim(SELDCBCREPORT_RS("BT")) = "" And Trim(SELDCBCREPORT_RS("PT")) = "" And Trim(SELDCBCREPORT_RS("PLATELET1")) = "" And Trim(SELDCBCREPORT_RS("PTTK")) = "" And Trim(SELDCBCREPORT_RS("TT")) = "" And Trim(SELDCBCREPORT_RS("FIBRINOGEN")) = "" And Trim(SELDCBCREPORT_RS("FDPDDIMER")) = "" And Trim(SELDCBCREPORT_RS("OTHERS")) = "" Then
                CrystalReport1.Formulas(11) = "CP = '" & "" & "'"
            Else
                CrystalReport1.Formulas(11) = "CP = '" & "ABC" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("NORMBOBLAS")) = "" And Trim(SELDCBCREPORT_RS("OTHERCELLR")) = "" And Trim(SELDCBCREPORT_RS("rbcmorpho")) = "" And Trim(SELDCBCREPORT_RS("reticcount")) = "" And Trim(SELDCBCREPORT_RS("mparasite")) = "" And Trim(SELDCBCREPORT_RS("ESR")) = "" Then
                CrystalReport1.Formulas(15) = "RBC = '" & "" & "'"
            Else
                CrystalReport1.Formulas(15) = "RBC = '" & "ABC" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("PMORPHOLOG")) = "" And Trim(SELDCBCREPORT_RS("FLD1")) = "" Then
                CrystalReport1.Formulas(16) = "PLATELETS = '" & "" & "'"
            Else
                CrystalReport1.Formulas(16) = "PLATELETS = '" & "ABC" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("PROLYMPHOC")) = "" And Trim(SELDCBCREPORT_RS("Sezarycell")) = "" And Trim(SELDCBCREPORT_RS("hairycells")) = "" And Trim(SELDCBCREPORT_RS("othercells")) = "" Then
                CrystalReport1.Formulas(17) = "prolyn = '" & "" & "'"
            Else
                CrystalReport1.Formulas(17) = "prolyn = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("blasts")) = "" And Trim(SELDCBCREPORT_RS("PROMYELOCY")) = "" And Trim(SELDCBCREPORT_RS("myelocytes")) = "" And Trim(SELDCBCREPORT_RS("metamyeloc")) = "" And Trim(SELDCBCREPORT_RS("plasmacell")) = "" And Trim(SELDCBCREPORT_RS("aTlymphEE")) = "" Then
                CrystalReport1.Formulas(18) = "blast = '" & "" & "'"
            Else
                CrystalReport1.Formulas(18) = "blast = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("neutrophi1")) = "" And Trim(SELDCBCREPORT_RS("lymphocyt1")) = "" And Trim(SELDCBCREPORT_RS("monocytE1")) = "" And Trim(SELDCBCREPORT_RS("eosinophi1")) = "" And Trim(SELDCBCREPORT_RS("basophill1")) = "" Then
                CrystalReport1.Formulas(19) = "cbcm = '" & "" & "'"
            Else
                CrystalReport1.Formulas(19) = "cbcm = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("HCT")) = "" And Trim(SELDCBCREPORT_RS("MCH")) = "" And Trim(SELDCBCREPORT_RS("MCV")) = "" And Trim(SELDCBCREPORT_RS("MCHC")) = "" And Trim(SELDCBCREPORT_RS("MPV")) = "" And Trim(SELDCBCREPORT_RS("RDW")) = "" And Trim(SELDCBCREPORT_RS("ABNORMAL")) = "" Then
                CrystalReport1.Formulas(20) = "rbci = '" & "" & "'"
            Else
                CrystalReport1.Formulas(20) = "rbci = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("neutrophil")) = "" And Trim(SELDCBCREPORT_RS("lymphocyte")) = "" And Trim(SELDCBCREPORT_RS("monocytes")) = "" And Trim(SELDCBCREPORT_RS("eosinophil")) = "" And Trim(SELDCBCREPORT_RS("basophils")) = "" And Trim(SELDCBCREPORT_RS("immgranulo")) = "" Then
                CrystalReport1.Formulas(21) = "dlc = '" & "" & "'"
            Else
                CrystalReport1.Formulas(21) = "dlc = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("HB")) = "" And Trim(SELDCBCREPORT_RS("RBC")) = "" And Trim(SELDCBCREPORT_RS("TLC")) = "" And Trim(SELDCBCREPORT_RS("platelets")) = "" Then
                CrystalReport1.Formulas(22) = "cbcc = '" & "" & "'"
            Else
                CrystalReport1.Formulas(22) = "cbcc = '" & "123" & "'"
            End If
End If
If opt_staff.Value = True Then
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
    If Trim(txt_modcode.Text) = "EK" Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISECS.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'           If Trim(prepath) = "" Then
'           CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'           Else
'           CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(SELRREQNORS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
    End If
End If
'***** STAFF
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport1.Action = 1

End Sub

Private Sub txt_inr_GotFocus()
TXT_PT_LostFocus (0)
End Sub

Private Sub txt_inr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'SendKeys vbTab
KeyAscii = 0
End If
End Sub

Private Sub txt_lymphocytes_LostFocus()
    txt_lymabs.Text = (TXT_TWBC.Text) * ((txt_lymphocytes.Text) / 100)
End Sub



Private Sub txt_monocytes_LostFocus()
    txt_mncyteabs.Text = (TXT_TWBC.Text) * ((txt_monocytes.Text) / 100)
End Sub

Private Sub txt_neutrophils_LostFocus()
    txt_ntphabs.Text = (TXT_TWBC.Text) * ((txt_neutrophils.Text) / 100)
End Sub


Private Sub TXT_PT_GotFocus(Index As Integer)
SELCPRNG_QR(0) = "PT"
'If Option3.Value = False Then
'    If Option1.Value = True Then
'    machname = "ACL 01"
'    ElseIf Option2.Value = True Then
'    machname = "ACL 02"
'    Else
    machname = cmb_machname.Text
'    End If
    Dim a As Boolean
    a = True
Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
    If SELCPRNG_RS.EOF = False Then
        While SELCPRNG_RS.EOF = False
        If Trim(SELCPRNG_RS("MACHINE")) = machname Then
        txt_ptctl.Text = SELCPRNG_RS("CONTROL")
        a = False
        Else
            If a = True Then
               txt_ptctl.Enabled = True
                txt_ptctl.Text = ""
              End If
        End If
        SELCPRNG_RS.MoveNext
        Wend
    End If
'End If
End Sub

Private Sub TXT_PT_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
'SendKeys vbTab
KeyAscii = 0
End If
End Sub

Private Sub TXT_PT_LostFocus(Index As Integer)
txt_inr.SetFocus
Dim VALE As String
If Not TXT_PT(0).Text = "" Then
    If InStr(1, TXT_PT(0), ">") Or InStr(1, TXT_PT(0), "<") Then
        VALE = Mid(TXT_PT(0), 2)
    Else
        VALE = TXT_PT(0).Text
    End If
    
    If Trim(txt_ptctl.Text) <> "" Then
        txt_inr.Text = Round((VALE / txt_ptctl.Text) ^ 1.12, 2)
    End If
End If
End Sub

Private Sub TXT_PTTK_GotFocus()
'SELCPRNG_QR(0) = "APTT"
'Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
'txt_apttctl.Text = SELCPRNG_RS("CONTROL")

SELCPRNG_QR(0) = "APTT"
'If Option3.Value = False Then
'    If Option1.Value = True Then
'    machname = "ACL 01"
'    ElseIf Option2.Value = True Then
'    machname = "ACL 02"
'    Else
    machname = cmb_machname.Text
'    End If
    Dim a As Boolean
    a = True
Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
    If SELCPRNG_RS.EOF = False Then
        While SELCPRNG_RS.EOF = False
        If Trim(SELCPRNG_RS("MACHINE")) = machname Then
        txt_apttctl.Text = SELCPRNG_RS("CONTROL")
        a = False
        Else
            If a = True Then
               txt_apttctl.Enabled = True
                txt_apttctl.Text = ""
              End If
        End If
        SELCPRNG_RS.MoveNext
        Wend
    End If
'End If
End Sub

Private Sub TXT_PTTK_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'SendKeys vbTab
KeyAscii = 0
End If
End Sub

Private Sub txt_rbcmorphology_GotFocus()
On Error GoTo m1:

w_rbcmorphology.Visible = False
Set d_rbcmorphology = w_rbcmorphology.ActiveDocument
d_rbcmorphology.Content = ""
d_rbcmorphology.Content = txt_rbcmorphology
w_rbcmorphology.Caption = "RBC MORPHOLOGY"
'w_frozen.WindowState = wdWindowStateMaximize
w_rbcmorphology.WindowState = wdWindowStateNormal
d_rbcmorphology.Content = txt_rbcmorphology
w_rbcmorphology.Visible = True
w_rbcmorphology.Activate
m1:
On Error Resume Next
Set w_rbcmorphology = CreateObject("Word.Application")
w_rbcmorphology.Documents.add ("C:\DISREPORTSEC\REPORT.DOT")
w_rbcmorphology.Caption = "RBC MORPHOLOGY"
'w_frozen.WindowState = wdWindowStateMaximize
w_rbcmorphology.WindowState = wdWindowStateNormal
Set d_rbcmorphology = w_rbcmorphology.ActiveDocument
d_rbcmorphology.Content = Trim(txt_rbcmorphology)
w_rbcmorphology.Visible = True
w_rbcmorphology.Activate
End Sub

Private Sub txt_rbcmorphology_LostFocus()
On Error GoTo m2:
Set d_rbcmorphology = w_rbcmorphology.ActiveDocument
txt_rbcmorphology.Text = d_rbcmorphology.Content
'flagimp = False
d_rbcmorphology.Content = ""
' Modified By Pankaj Baggal 03/09/2004
'd_rbcmorphology.SaveAs "C:\DISREPORTSEC\fronzen.DOC"
d_rbcmorphology.Close (wdDoNotSaveChanges)
w_rbcmorphology.quit
'W1.Visible = False
Exit Sub
m2:
MsgBox Err.Description

End Sub

Public Sub INTERF()
        selhinterface_qr(0) = req_chglen(cmb_reqno.Text)
        Set selhinterface_rs = selhinterface_qr.OpenResultset(1, 1)
        If selhinterface_rs.RowCount > 0 Then
         TXT_HB = selhinterface_rs("hb")
         TXT_RBC = selhinterface_rs("rbc")
         TXT_TWBC = selhinterface_rs("tlc")
         txt_platelets = selhinterface_rs("platelets")
         txt_neutrophils = selhinterface_rs("neutrophil")
         txt_lymphocytes = selhinterface_rs("lymphocyte")
         txt_monocytes = selhinterface_rs("monocytes")
         txt_eosinophils = selhinterface_rs("eosinophil")
         txt_basophils = selhinterface_rs("basophils")
         txt_immaturegranulocytes = selhinterface_rs("immgranulo")
         TXT_HCT = selhinterface_rs("hct")
         TXT_MCH = selhinterface_rs("mch")
         TXT_MCHC = selhinterface_rs("mchc")
         TXT_MCV = selhinterface_rs("mcv")
         TXT_MPV = selhinterface_rs("mpv")
         TXT_RDW = selhinterface_rs("rdw")
        End If
        selcinterface_qr(0) = req_chglen(cmb_reqno.Text)
        Set selcinterface_rs = selcinterface_qr.OpenResultset(1, 1)
        Do While selcinterface_rs.EOF = False
          If selcinterface_rs("testcode") = "U708" Then
          TXT_PT(0).Text = selcinterface_rs("TESTRESULT")
           
          ElseIf selcinterface_rs("testcode") = "U710" Then
            TXT_PTTK = selcinterface_rs("TESTRESULT")
          ElseIf selcinterface_rs("testcode") = "U714" Then
            TXT_FDPDMIER = selcinterface_rs("TESTRESULT")
            
          ElseIf selcinterface_rs("testcode") = "U715" Then
            TXT_FIBRINOGEN = selcinterface_rs("TESTRESULT")
          End If
          
          selcinterface_rs.MoveNext
          'selcinterface_rs.MoveNext
          
        Loop
        
        
        
'' hemato  interface fetching ends

End Sub

Public Sub HBAGE()

If (txt_age.Text <= 0.2 And txt_age.Text >= "0.00") And Trim(txt_sex.Text) = "M" Then
    HAGE = 0.2
    HBSEX = "M"
End If

If (txt_age.Text <= 0.2 And txt_age.Text >= "0.00") And Trim(txt_sex.Text) = "F" Then
    HAGE = 0.2
    HBSEX = "F"
End If

If (txt_age.Text <= 0.1 And txt_age.Text >= 0.03) And Trim(txt_sex.Text) = "M" Then
    HAGE = 0.1
    HBSEX = "M"
End If

If (txt_age.Text <= 0.1 And txt_age.Text >= 0.03) And Trim(txt_sex.Text) = "F" Then
    HAGE = 0.1
    HBSEX = "F"
End If

If (txt_age.Text <= 0.2 And txt_age.Text > 0.1) And Trim(txt_sex.Text) = "M" Then
    HAGE = 0.2
    HBSEX = "M"
End If

If (txt_age.Text <= 0.2 And txt_age.Text > 0.1) And Trim(txt_sex.Text) = "F" Then
    HAGE = 0.2
    HBSEX = "F"
End If

If (txt_age.Text <= 0.3 And txt_age.Text > 0.2) And Trim(txt_sex.Text) = "M" Then
    HAGE = 0.3
    HBSEX = "M"
End If

If (txt_age.Text <= 0.3 And txt_age.Text > 0.2) And Trim(txt_sex.Text) = "F" Then
    HAGE = 0.3
    HBSEX = "F"
End If

If (txt_age.Text <= 0.6 And txt_age.Text > 0.3) And Trim(txt_sex.Text) = "M" Then
    HAGE = 0.6
    HBSEX = "M"
End If

If (txt_age.Text <= 0.6 And txt_age.Text > 0.3) And Trim(txt_sex.Text) = "F" Then
    HAGE = 0.6
    HBSEX = "F"
End If

If (txt_age.Text <= 2 And txt_age.Text > 0.6) And Trim(txt_sex.Text) = "M" Then
    HAGE = 2
    HBSEX = "M"
End If

If (txt_age.Text <= 2 And txt_age.Text > 0.6) And Trim(txt_sex.Text) = "F" Then
    HAGE = 2
    HBSEX = "F"
End If

If (txt_age.Text <= 6 And txt_age.Text > 2) And Trim(txt_sex.Text) = "M" Then
    HAGE = 6
    HBSEX = "M"
End If

If (txt_age.Text <= 6 And txt_age.Text > 2) And Trim(txt_sex.Text) = "F" Then
    HAGE = 6
    HBSEX = "F"
End If

If (txt_age.Text <= 12 And txt_age.Text > 6) And Trim(txt_sex.Text) = "M" Then
    HAGE = 12
    HBSEX = "M"
End If

If (txt_age.Text <= 12 And txt_age.Text > 6) And Trim(txt_sex.Text) = "F" Then
    HAGE = 12
    HBSEX = "F"
End If

If (txt_age.Text <= 106 And txt_age.Text > 12) And Trim(txt_sex.Text) = "M" Then
    HAGE = 106
    HBSEX = "M"
End If

If (txt_age.Text <= 106 And txt_age.Text > 12) And Trim(txt_sex.Text) = "F" Then
    HAGE = 106
    HBSEX = "F"
End If

'If (txt_age.Text <= 120 And txt_age.Text > 12) And Trim(txt_sex.Text) = "M" Then
'    HAGE = 120
'    HBSEX = "M"
'End If
'
'If (txt_age.Text <= 120 And txt_age.Text > 12) And Trim(txt_sex.Text) = "F" Then
'    HAGE = 120
'    HBSEX = "F"
'End If




End Sub
Public Sub inshemcbc()
    For I = 0 To 5
          If I = 0 Then
            INSCP_QR(0) = req_chglen(cmb_reqno.Text)
            INSCP_QR(1) = "PT"
            
            SELCPRNG_QR(0) = "PT"
            Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
            If Len(Trim(TXT_PT(0).Text)) > 0 Then
                INSCP_QR(2) = Trim(TXT_PT(0).Text)
                
                INSCP_QR(3) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(TXT_PT(0).Text))
'                If Option3.Value = True Then
'                INSCP_QR(4) = txt_ptctl.Text
'                Else
              '  INSCP_QR(4) = cmb_machname.Text
'                End If
            Else
                INSCP_QR(2) = ""
                INSCP_QR(3) = ""
                INSCP_QR(4) = ""
            End If
               INSCP_QR(4) = txt_ptctl.Text
            INSCP_QR(5) = cmb_enterby.Text
            INSCP_QR(6) = SELCPRNG_RS("UOM")
            INSCP_QR(7) = SELCPRNG_RS("RANGE")
'            If Option1.Value = True Then
'
'            INSCP_QR(8) = "ACL 100"
'            ElseIf Option2.Value = True Then
'            INSCP_QR(8) = "ACL 200"
'            ElseIf Option3.Value = True Then
'            INSCP_QR(8) = "ELP 01"
'            Else
            INSCP_QR(8) = cmb_machname.Text
'            End If
            'INSCP_QR(9) = ""
            'INSCP_QR(10) =
            INSCP_QR.Execute
      
        End If
       If I = 1 Then
            INSCP_QR(0) = req_chglen(cmb_reqno.Text)
            INSCP_QR(1) = "INR"
            SELCPRNG_QR(0) = "INR"
            Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
            If Len(Trim(txt_inr.Text)) > 0 Then
                INSCP_QR(2) = Trim(txt_inr.Text)
                
                INSCP_QR(3) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(txt_inr.Text))
            Else
                INSCP_QR(2) = ""
                INSCP_QR(3) = ""
            End If
            INSCP_QR(4) = ""
            INSCP_QR(5) = cmb_enterby.Text
            INSCP_QR(6) = SELCPRNG_RS("UOM")
            INSCP_QR(7) = SELCPRNG_RS("RANGE")
'            If Option1.Value = True Then
'
'            INSCP_QR(8) = "ACL 100"
'            ElseIf Option2.Value = True Then
'            INSCP_QR(8) = "ACL 200"
'            ElseIf Option3.Value = True Then
'            INSCP_QR(8) = "ELP 01"
'            Else
            INSCP_QR(8) = cmb_machname.Text
'            End If
            INSCP_QR.Execute
       End If
       If I = 2 Then
            INSCP_QR(0) = req_chglen(cmb_reqno.Text)
            INSCP_QR(1) = "APTT"
            SELCPRNG_QR(0) = "APTT"
                Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
            If Len(Trim(TXT_PTTK.Text)) > 0 Then
                INSCP_QR(2) = Trim(TXT_PTTK.Text)
                
                INSCP_QR(3) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(TXT_PTTK.Text))
                'INSCP_QR(4) = SELCPRNG_RS("control")
'                If Option3.Value = True Then
'                INSCP_QR(4) = txt_apttctl.Text
'                Else
                INSCP_QR(4) = Trim(txt_ttctl.Text)
'                End If
            Else
                INSCP_QR(2) = ""
                INSCP_QR(3) = ""
            End If
            
            INSCP_QR(5) = cmb_enterby.Text
            INSCP_QR(6) = SELCPRNG_RS("UOM")
            INSCP_QR(7) = SELCPRNG_RS("RANGE")
'            If Option1.Value = True Then
'
'            INSCP_QR(8) = "ACL 100"
'            ElseIf Option2.Value = True Then
'            INSCP_QR(8) = "ACL 200"
'            ElseIf Option3.Value = True Then
'            INSCP_QR(8) = "ELP 01"
'            Else
            INSCP_QR(8) = cmb_machname.Text
'            End If
            INSCP_QR.Execute
       End If
       If I = 3 Then
            INSCP_QR(0) = req_chglen(cmb_reqno.Text)
            INSCP_QR(1) = "TT"
            SELCPRNG_QR(0) = "TT"
                Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
            If Len(Trim(TXT_TT.Text)) > 0 Then
                INSCP_QR(2) = Trim(TXT_TT.Text)
                
                INSCP_QR(3) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(TXT_TT.Text))
    
                INSCP_QR(4) = Trim(txt_ttctl.Text) ''MANUAL CONTROL ADD
            Else
                INSCP_QR(2) = ""
                INSCP_QR(3) = ""
                INSCP_QR(4) = ""
            End If
            INSCP_QR(5) = cmb_enterby.Text
            INSCP_QR(6) = SELCPRNG_RS("UOM")
            INSCP_QR(7) = SELCPRNG_RS("RANGE")
'            If Option1.Value = True Then
'
'            INSCP_QR(8) = "ACL 100"
'            ElseIf Option2.Value = True Then
'            INSCP_QR(8) = "ACL 200"
'            ElseIf Option3.Value = True Then
'            INSCP_QR(8) = "ELP 01"
'            Else
            INSCP_QR(8) = cmb_machname.Text
'            End If
            INSCP_QR.Execute
       End If
       If I = 4 Then
            INSCP_QR(0) = req_chglen(cmb_reqno.Text)
            INSCP_QR(1) = "FIBRINOGEN"
            SELCPRNG_QR(0) = "FIBRINOGEN"
                Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
            If Len(Trim(TXT_FIBRINOGEN.Text)) > 0 Then
                INSCP_QR(2) = Trim(TXT_FIBRINOGEN.Text)
                
                INSCP_QR(3) = RNGFLG(SELCPRNG_RS("RANGE"), Trim(TXT_FIBRINOGEN.Text))
            Else
                INSCP_QR(2) = ""
                INSCP_QR(3) = ""
            End If
            INSCP_QR(4) = "" ''NO CONTROL ADD
            INSCP_QR(5) = cmb_enterby.Text
            INSCP_QR(6) = SELCPRNG_RS("UOM")
            INSCP_QR(7) = SELCPRNG_RS("RANGE")
'            If Option1.Value = True Then
'
'            INSCP_QR(8) = "ACL 100"
'            ElseIf Option2.Value = True Then
'            INSCP_QR(8) = "ACL 200"
'            ElseIf Option3.Value = True Then
            INSCP_QR(8) = cmb_machname.Text
'            End If
            INSCP_QR.Execute
       End If
       
       If I = 5 Then
            Dim NRANGE As String
            Dim DDIMER As Integer
            INSCP_QR(0) = req_chglen(cmb_reqno.Text)
            INSCP_QR(1) = "D-DIMER"
            SELCPRNG_QR(0) = "D-DIMER"
            Set SELCPRNG_RS = SELCPRNG_QR.OpenResultset(2, 1)
            If Len(Trim(TXT_FDPDMIER.Text)) > 0 Then
                INSCP_QR(2) = Trim(TXT_FDPDMIER.Text)
                'DDIMER = Val(Mid(Trim(SELCPRNG_RS("RANGE")), 2, 3))
                DDIMER = Val(Mid(Trim(SELCPRNG_RS("RANGE")), 2, 3))
                'ddimer = 200
                NRANGE = Trim(TXT_FDPDMIER.Text)
                
                '***** Shraddha 02/05/2011 *****
                 
                If Val(Mid(NRANGE, 2, 4)) < DDIMER Then
                    INSCP_QR(3) = "N"
                ElseIf Val(Mid(NRANGE, 2, 4)) >= DDIMER Then
                    INSCP_QR(3) = "H"
                End If
                
                '********* 02/05/2011 **********
                
                '******* Original ***************
                
'                If InStr(1, NRANGE, "<") > 0 Then
'                    Do While InStr(NRANGE, "<") > 0
'                        NRANGE = Replace(NRANGE, "<", "*")
'                    Loop
'                End If
'                If InStr(1, NRANGE, ">") > 0 Then
'                Do While InStr(NRANGE, ">") > 0
'                        NRANGE = Replace(NRANGE, ">", "*")
'                    Loop
'                End If
'
'                '***** 17/02/10 for 4 digit comparision *****
'                If Trim(Mid(NRANGE, 5, 1)) <> "*" Then
'
'                    If Val(Mid(NRANGE, 2, 4)) < DDIMER Then
'                        INSCP_QR(3) = "L"
'                    ElseIf Val(Mid(NRANGE, 2, 4)) > DDIMER Then
'                        INSCP_QR(3) = "H"
'                    ElseIf Val(Mid(NRANGE, 2, 4)) = DDIMER Then
'                        INSCP_QR(3) = "N"
'                    End If
'                Else
'                    If Val(Mid(NRANGE, 2, 3)) < DDIMER Then
'                '******************************************************
'                        INSCP_QR(3) = "L"
'                    ElseIf Val(Mid(NRANGE, 2, 3)) > DDIMER Then
'                        INSCP_QR(3) = "H"
'                    ElseIf Val(Mid(NRANGE, 2, 3)) = DDIMER Then
'                        INSCP_QR(3) = "N"
'                    End If
'                End If
'                'INSCP_QR(3) = Mid(SELCPRNG_RS("RANGE"), 2)
                '************ Original ********************
            Else
                INSCP_QR(2) = ""
                INSCP_QR(3) = ""
            End If
            INSCP_QR(4) = "" ''NO CONTROL ADD
            INSCP_QR(5) = cmb_enterby.Text
            INSCP_QR(6) = SELCPRNG_RS("UOM")
            INSCP_QR(7) = SELCPRNG_RS("RANGE")
'            If Option1.Value = True Then
'            INSCP_QR(8) = "ACL 100"
'            ElseIf Option2.Value = True Then
'            INSCP_QR(8) = "ACL 200"
'            ElseIf Option3.Value = True Then
'            INSCP_QR(8) = "ELP 01"
'            Else
            INSCP_QR(8) = cmb_machname.Text
'            End If
            INSCP_QR.Execute
       End If
    Next
End Sub

Public Sub CBCDETAIL()
        'For i = 1 To 16
            
            SELHINTFVAL_QR(0) = req_chglen(cmb_reqno.Text)
            Set SELHINTFVAL_RS = SELHINTFVAL_QR.OpenResultset(1, 1)
                        
            INSDHEMCBCDET_QR(0) = req_chglen(cmb_reqno.Text)
                            
                        
            If Len(Trim(TXT_HB.Text)) > 0 Then
                
                    INSDHEMCBCDET_QR(1) = "HAEMOGLOBIN"
                    INSDHEMCBCDET_QR(2) = Trim(TXT_HB.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "HB"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "HB"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(TXT_HB.Text)) <= Val(HBHV) And Val(Trim(TXT_HB.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(TXT_HB.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(TXT_HB.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
            End If
            
            If Len(Trim(TXT_RBC.Text)) > 0 Then
                
                    INSDHEMCBCDET_QR(1) = "RBC COUNT"
                    INSDHEMCBCDET_QR(2) = Trim(TXT_RBC.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "RBC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "RBC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(TXT_RBC.Text)) <= Val(HBHV) And Val(Trim(TXT_RBC.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(TXT_RBC.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(TXT_RBC.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
                
            End If

            If Len(Trim(TXT_HCT.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "HAEMATOCRIT"
                    INSDHEMCBCDET_QR(2) = Trim(TXT_HCT.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "HCT"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "HCT"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    
                    If Val(Trim(TXT_HCT.Text)) <= Val(HBHV) And Val(Trim(TXT_HCT.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(TXT_HCT.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(TXT_HCT.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
            
            End If
            
            If Len(Trim(TXT_MCV.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "MCV"
                    INSDHEMCBCDET_QR(2) = Trim(TXT_MCV.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "MCV"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "MCV"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    
                    If Val(Trim(TXT_MCV.Text)) <= Val(HBHV) And Val(Trim(TXT_MCV.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(TXT_MCV.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(TXT_MCV.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
            
            End If
            
            If Len(Trim(TXT_MCH.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "MCH"
                    INSDHEMCBCDET_QR(2) = Trim(TXT_MCH.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "MCH"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "MCH"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    
                    If Val(Trim(TXT_MCH.Text)) <= Val(HBHV) And Val(Trim(TXT_MCH.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(TXT_MCH.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(TXT_MCH.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
            
            End If
            
            If Len(Trim(TXT_MCHC.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "MCHC"
                    INSDHEMCBCDET_QR(2) = Trim(TXT_MCHC.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "MCHC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "MCHC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    
                    If Val(Trim(TXT_MCHC.Text)) <= Val(HBHV) And Val(Trim(TXT_MCHC.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(TXT_MCHC.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(TXT_MCHC.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
           
            End If
            
            If Len(Trim(TXT_RDW.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "RDW"
                    INSDHEMCBCDET_QR(2) = Trim(TXT_RDW.Text)
                    'HBAGE
                    RDWAGE
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "RDW"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "RDW"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    
                    If Val(Trim(TXT_RDW.Text)) <= Val(HBHV) And Val(Trim(TXT_RDW.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(TXT_RDW.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(TXT_RDW.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
            
            End If

            If Len(Trim(txt_platelets.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "PLATELETS"
                    INSDHEMCBCDET_QR(2) = Trim(txt_platelets.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Platelets"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Platelets"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_platelets.Text)) <= Val(HBHV) And Val(Trim(txt_platelets.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(txt_platelets.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(txt_platelets.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
                
            End If
            
            If Len(Trim(TXT_MPV.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "MPV"
                    INSDHEMCBCDET_QR(2) = Trim(TXT_MPV.Text)
                    'HBAGE
                    MPVAGE
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "MPV"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "MPV"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    
                    If Val(Trim(TXT_MPV.Text)) <= Val(HBHV) And Val(Trim(TXT_MPV.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(TXT_MPV.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(TXT_MPV.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                

            
            End If

            If Len(Trim(TXT_TWBC.Text)) > 0 Then
                
                    INSDHEMCBCDET_QR(1) = "WBC COUNT"
                    INSDHEMCBCDET_QR(2) = Trim(TXT_TWBC.Text)
                    HBAGE
                    'TLCAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "TLC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "TLC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(TXT_TWBC.Text)) <= Val(HBHV) And Val(Trim(TXT_TWBC.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(TXT_TWBC.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(TXT_TWBC.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
            End If

            If Len(Trim(txt_neutrophils.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "NEUTROPHIL %"
                    INSDHEMCBCDET_QR(2) = Trim(txt_neutrophils.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "neutrophil %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "neutrophil %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_neutrophils.Text)) <= Val(HBHV) And Val(Trim(txt_neutrophils.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(txt_neutrophils.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(txt_neutrophils.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
                
            End If
            
            If Len(Trim(txt_lymphocytes.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "LYMPHOCYTE %"
                    INSDHEMCBCDET_QR(2) = Trim(txt_lymphocytes.Text)
                    'HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Lymphocyte %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Lymphocyte %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_lymphocytes.Text)) <= Val(HBHV) And Val(Trim(txt_lymphocytes.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(txt_lymphocytes.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(txt_lymphocytes.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
                
            End If
            
            If Len(Trim(txt_monocytes.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "MONOCYTE %"
                    INSDHEMCBCDET_QR(2) = Trim(txt_monocytes.Text)
                    'HBAGE
                    MONOCYTEAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Monocyte %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Monocyte %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_monocytes.Text)) <= Val(HBHV) And Val(Trim(txt_monocytes.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(txt_monocytes.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(txt_monocytes.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
              
            End If
            
            If Len(Trim(txt_eosinophils.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "EOSINOPHIL %"
                    INSDHEMCBCDET_QR(2) = Trim(txt_eosinophils.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "eosinophils %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "eosinophils %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_eosinophils.Text)) <= Val(HBHV) And Val(Trim(txt_eosinophils.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(txt_eosinophils.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(txt_eosinophils.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
                
            End If
            
            If Len(Trim(txt_basophils.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "BASOPHIL %"
                    INSDHEMCBCDET_QR(2) = Trim(txt_basophils.Text)
                    'HBAGE
                    BASOPHILSAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Basophils %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Basophils %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_basophils.Text)) <= Val(HBHV) And Val(Trim(txt_basophils.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(txt_basophils.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(txt_basophils.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
                
            
            End If

            If Len(Trim(txt_ntphabs.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "NEUTROPHIL ABS"
                    INSDHEMCBCDET_QR(2) = Trim(txt_ntphabs.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Neutrophils abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Neutrophils abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_ntphabs.Text)) <= Val(HBHV) And Val(Trim(txt_ntphabs.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(txt_ntphabs.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(txt_ntphabs.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
            End If
            
                        
            If Len(Trim(txt_lymabs.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "LYMPHOCYTE ABS"
                    INSDHEMCBCDET_QR(2) = Trim(txt_lymabs.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Lymphocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Lymphocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_lymabs.Text)) <= Val(HBHV) And Val(Trim(txt_lymabs.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(txt_lymabs.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(txt_lymabs.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
            End If

                        
            If Len(Trim(txt_mncyteabs.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "MONOCYTE ABS"
                    INSDHEMCBCDET_QR(2) = Trim(txt_mncyteabs.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Monocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Monocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_mncyteabs.Text)) <= Val(HBHV) And Val(Trim(txt_mncyteabs.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(txt_mncyteabs.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(txt_mncyteabs.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
            End If

            
            
            If Len(Trim(txt_esnabs.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "EOSINOPHIL ABS"
                    INSDHEMCBCDET_QR(2) = Trim(txt_esnabs.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Monocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Monocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_esnabs.Text)) <= Val(HBHV) And Val(Trim(txt_esnabs.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(txt_esnabs.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(txt_esnabs.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
            End If

            
                        
            If Len(Trim(txt_bsphabs.Text)) > 0 Then
                
                
                    INSDHEMCBCDET_QR(1) = "BASOPHIL ABS"
                    INSDHEMCBCDET_QR(2) = Trim(txt_bsphabs.Text)
                    'HBAGE
                    BASOPHILSAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Basophils abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Basophils abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_bsphabs.Text)) <= Val(HBHV) And Val(Trim(txt_bsphabs.Text)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(txt_bsphabs.Text)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(txt_bsphabs.Text)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If
                        
                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""
                        
                        INSDHEMCBCDET_QR.Execute
                        
            End If
            

End Sub

Public Sub TLCAGE()

If (txt_age.Text <= 120 And txt_age.Text >= 12) And Trim(txt_sex.Text) = "M" Then
    HAGE = 120
    HBSEX = "M"
End If

If (txt_age.Text <= 120 And txt_age.Text >= 12) And Trim(txt_sex.Text) = "F" Then
    HAGE = 120
    HBSEX = "F"
End If

End Sub

Public Sub RDWAGE()

If (txt_age.Text <= 106 And txt_age.Text >= 0) And Trim(txt_sex.Text) = "M" Then
    HAGE = 106
    HBSEX = "M"
End If

If (txt_age.Text <= 106 And txt_age.Text >= 0) And Trim(txt_sex.Text) = "F" Then
    HAGE = 106
    HBSEX = "F"
End If

End Sub

Public Sub MPVAGE()

If (txt_age.Text <= 106 And txt_age.Text >= 0) And Trim(txt_sex.Text) = "M" Then
    HAGE = 106
    HBSEX = "M"
End If

If (txt_age.Text <= 106 And txt_age.Text >= 0) And Trim(txt_sex.Text) = "F" Then
    HAGE = 106
    HBSEX = "F"
End If

End Sub

Public Sub MONOCYTEAGE()

If (txt_age.Text <= 0.2 And txt_age.Text >= "0.00") And Trim(txt_sex.Text) = "M" Then
    HAGE = 0.2
    HBSEX = "M"
End If

If (txt_age.Text <= 0.2 And txt_age.Text >= "0.00") And Trim(txt_sex.Text) = "F" Then
    HAGE = 0.2
    HBSEX = "F"
End If

If (txt_age.Text <= 0.1 And txt_age.Text >= 0.03) And Trim(txt_sex.Text) = "M" Then
    HAGE = 0.1
    HBSEX = "M"
End If

If (txt_age.Text <= 0.1 And txt_age.Text >= 0.03) And Trim(txt_sex.Text) = "F" Then
    HAGE = 0.1
    HBSEX = "F"
End If

If (txt_age.Text <= 0.3 And txt_age.Text > 0.1) And Trim(txt_sex.Text) = "M" Then
    HAGE = 0.3
    HBSEX = "M"
End If

If (txt_age.Text <= 0.3 And txt_age.Text > 0.1) And Trim(txt_sex.Text) = "F" Then
    HAGE = 0.3
    HBSEX = "F"
End If

If (txt_age.Text <= 0.6 And txt_age.Text > 0.3) And Trim(txt_sex.Text) = "M" Then
    HAGE = 0.6
    HBSEX = "M"
End If

If (txt_age.Text <= 0.6 And txt_age.Text > 0.3) And Trim(txt_sex.Text) = "F" Then
    HAGE = 0.6
    HBSEX = "F"
End If

If (txt_age.Text <= 2 And txt_age.Text > 0.6) And Trim(txt_sex.Text) = "M" Then
    HAGE = 2
    HBSEX = "M"
End If

If (txt_age.Text <= 2 And txt_age.Text > 0.6) And Trim(txt_sex.Text) = "F" Then
    HAGE = 2
    HBSEX = "F"
End If

If (txt_age.Text <= 6 And txt_age.Text > 2) And Trim(txt_sex.Text) = "M" Then
    HAGE = 6
    HBSEX = "M"
End If

If (txt_age.Text <= 6 And txt_age.Text > 2) And Trim(txt_sex.Text) = "F" Then
    HAGE = 6
    HBSEX = "F"
End If

If (txt_age.Text <= 12 And txt_age.Text > 6) And Trim(txt_sex.Text) = "M" Then
    HAGE = 12
    HBSEX = "M"
End If

If (txt_age.Text <= 12 And txt_age.Text > 6) And Trim(txt_sex.Text) = "F" Then
    HAGE = 12
    HBSEX = "F"
End If

If (txt_age.Text <= 106 And txt_age.Text > 12) And Trim(txt_sex.Text) = "M" Then
    HAGE = 106
    HBSEX = "M"
End If

If (txt_age.Text <= 106 And txt_age.Text > 12) And Trim(txt_sex.Text) = "F" Then
    HAGE = 106
    HBSEX = "F"
End If

End Sub

Public Sub BASOPHILSAGE()

If (txt_age.Text <= 1 And txt_age.Text >= "0.00") And Trim(txt_sex.Text) = "M" Then
    HAGE = 1
    HBSEX = "M"
End If

If (txt_age.Text <= 1 And txt_age.Text >= "0.00") And Trim(txt_sex.Text) = "F" Then
    HAGE = 1
    HBSEX = "F"
End If


If (txt_age.Text <= 4 And txt_age.Text > 1) And Trim(txt_sex.Text) = "M" Then
    HAGE = 4
    HBSEX = "M"
End If

If (txt_age.Text <= 4 And txt_age.Text > 1) And Trim(txt_sex.Text) = "F" Then
    HAGE = 4
    HBSEX = "F"
End If

If (txt_age.Text <= 6 And txt_age.Text > 4) And Trim(txt_sex.Text) = "M" Then
    HAGE = 6
    HBSEX = "M"
End If

If (txt_age.Text <= 6 And txt_age.Text > 4) And Trim(txt_sex.Text) = "F" Then
    HAGE = 6
    HBSEX = "F"
End If

If (txt_age.Text <= 10 And txt_age.Text > 6) And Trim(txt_sex.Text) = "M" Then
    HAGE = 10
    HBSEX = "M"
End If

If (txt_age.Text <= 10 And txt_age.Text > 6) And Trim(txt_sex.Text) = "F" Then
    HAGE = 10
    HBSEX = "F"
End If

If (txt_age.Text <= 21 And txt_age.Text > 10) And Trim(txt_sex.Text) = "M" Then
    HAGE = 21
    HBSEX = "M"
End If

If (txt_age.Text <= 21 And txt_age.Text > 10) And Trim(txt_sex.Text) = "F" Then
    HAGE = 21
    HBSEX = "F"
End If

If (txt_age.Text <= 106 And txt_age.Text > 21) And Trim(txt_sex.Text) = "M" Then
    HAGE = 106
    HBSEX = "M"
End If

If (txt_age.Text <= 106 And txt_age.Text > 21) And Trim(txt_sex.Text) = "F" Then
    HAGE = 106
    HBSEX = "F"
End If

End Sub

Public Sub CBCDTLUPDT()
    
            
                        
            
            
            
            
                        
            If Len(Trim(TXT_HB.Text)) > 0 Then
                
                    
                    UPDDHEMCBCDET_QR(0) = Trim(TXT_HB.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "HB"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "HB"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(TXT_HB.Text)) <= Val(HBHV) And Val(Trim(TXT_HB.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(TXT_HB.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(TXT_HB.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "HAEMOGLOBIN"
                        UPDDHEMCBCDET_QR.Execute
                        
                
            End If
            
            If Len(Trim(TXT_RBC.Text)) > 0 Then
                
                    
                    UPDDHEMCBCDET_QR(0) = Trim(TXT_RBC.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "RBC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "RBC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(TXT_RBC.Text)) <= Val(HBHV) And Val(Trim(TXT_RBC.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(TXT_RBC.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(TXT_RBC.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "RBC COUNT"
                        UPDDHEMCBCDET_QR.Execute
                        
                
                
            End If

            If Len(Trim(TXT_HCT.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(TXT_HCT.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "HCT"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "HCT"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    
                    If Val(Trim(TXT_HCT.Text)) <= Val(HBHV) And Val(Trim(TXT_HCT.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(TXT_HCT.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(TXT_HCT.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "HAEMATOCRIT"
                        UPDDHEMCBCDET_QR.Execute
                        
                
            
            End If
            
            If Len(Trim(TXT_MCV.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(TXT_MCV.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "MCV"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "MCV"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    
                    If Val(Trim(TXT_MCV.Text)) <= Val(HBHV) And Val(Trim(TXT_MCV.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(TXT_MCV.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(TXT_MCV.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "MCV"
                        UPDDHEMCBCDET_QR.Execute
                        
                
            
            End If
            
            If Len(Trim(TXT_MCH.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(TXT_MCH.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "MCH"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "MCH"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    
                    If Val(Trim(TXT_MCH.Text)) <= Val(HBHV) And Val(Trim(TXT_MCH.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(TXT_MCH.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(TXT_MCH.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "MCH"
                        UPDDHEMCBCDET_QR.Execute
                        
                
            
            End If
            
            If Len(Trim(TXT_MCHC.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(TXT_MCHC.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "MCHC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "MCHC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    
                    If Val(Trim(TXT_MCHC.Text)) <= Val(HBHV) And Val(Trim(TXT_MCHC.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(TXT_MCHC.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(TXT_MCHC.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "MCHC"
                        UPDDHEMCBCDET_QR.Execute
                        
                
           
            End If
            
            If Len(Trim(TXT_RDW.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(TXT_RDW.Text)
                    'HBAGE
                    RDWAGE
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "RDW"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "RDW"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    
                    If Val(Trim(TXT_RDW.Text)) <= Val(HBHV) And Val(Trim(TXT_RDW.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(TXT_RDW.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(TXT_RDW.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "RDW"
                        UPDDHEMCBCDET_QR.Execute
                        
                
            
            End If

            If Len(Trim(txt_platelets.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(txt_platelets.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Platelets"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Platelets"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_platelets.Text)) <= Val(HBHV) And Val(Trim(txt_platelets.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(txt_platelets.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(txt_platelets.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "PLATELETS"
                        UPDDHEMCBCDET_QR.Execute
                        
                
                
            End If
            
            If Len(Trim(TXT_MPV.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(TXT_MPV.Text)
                    'HBAGE
                    MPVAGE
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "MPV"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "MPV"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    
                    If Val(Trim(TXT_MPV.Text)) <= Val(HBHV) And Val(Trim(TXT_MPV.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(TXT_MPV.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(TXT_MPV.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "MPV"
                        UPDDHEMCBCDET_QR.Execute
                        
                

            
            End If

            If Len(Trim(TXT_TWBC.Text)) > 0 Then
                
                    UPDDHEMCBCDET_QR(0) = Trim(TXT_TWBC.Text)
                    HBAGE
                    'TLCAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "TLC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "TLC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(TXT_TWBC.Text)) <= Val(HBHV) And Val(Trim(TXT_TWBC.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(TXT_TWBC.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(TXT_TWBC.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "WBC COUNT"
                        UPDDHEMCBCDET_QR.Execute
                        
                
            End If

            If Len(Trim(txt_neutrophils.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(txt_neutrophils.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "neutrophil %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "neutrophil %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_neutrophils.Text)) <= Val(HBHV) And Val(Trim(txt_neutrophils.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(txt_neutrophils.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(txt_neutrophils.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "NEUTROPHIL %"
                        UPDDHEMCBCDET_QR.Execute
                        
                
                
            End If
            
            If Len(Trim(txt_lymphocytes.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(txt_lymphocytes.Text)
                    'HBAGE
                    LYMPHPCNTAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Lymphocyte %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Lymphocyte %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_lymphocytes.Text)) <= Val(HBHV) And Val(Trim(txt_lymphocytes.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(txt_lymphocytes.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(txt_lymphocytes.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "LYMPHOCYTE %"
                        UPDDHEMCBCDET_QR.Execute
                        
                
                
            End If
            
            If Len(Trim(txt_monocytes.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(txt_monocytes.Text)
                    'HBAGE
                    MONOCYTEAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Monocyte %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Monocyte %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_monocytes.Text)) <= Val(HBHV) And Val(Trim(txt_monocytes.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(txt_monocytes.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(txt_monocytes.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "MONOCYTE %"
                        UPDDHEMCBCDET_QR.Execute
                        
                
              
            End If
            
            If Len(Trim(txt_eosinophils.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(txt_eosinophils.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "eosinophils %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "eosinophils %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_eosinophils.Text)) <= Val(HBHV) And Val(Trim(txt_eosinophils.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(txt_eosinophils.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(txt_eosinophils.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "EOSINOPHIL %"
                        UPDDHEMCBCDET_QR.Execute
                        
                
                
            End If
            
            If Len(Trim(txt_basophils.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(txt_basophils.Text)
                    'HBAGE
                    BASOPHILSAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Basophils %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Basophils %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_basophils.Text)) <= Val(HBHV) And Val(Trim(txt_basophils.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(txt_basophils.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(txt_basophils.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "BASOPHIL %"
                        UPDDHEMCBCDET_QR.Execute
                        
                
            
            End If

            If Len(Trim(txt_ntphabs.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(txt_ntphabs.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Neutrophils abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Neutrophils abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_ntphabs.Text)) <= Val(HBHV) And Val(Trim(txt_ntphabs.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(txt_ntphabs.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(txt_ntphabs.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "NEUTROPHIL ABS"
                        UPDDHEMCBCDET_QR.Execute
                        
            End If
            
                        
            If Len(Trim(txt_lymabs.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(txt_lymabs.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Lymphocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Lymphocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_lymabs.Text)) <= Val(HBHV) And Val(Trim(txt_lymabs.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(txt_lymabs.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(txt_lymabs.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "LYMPHOCYTE ABS"
                        UPDDHEMCBCDET_QR.Execute
                        
            End If

                        
            If Len(Trim(txt_mncyteabs.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(txt_mncyteabs.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Monocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Monocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_mncyteabs.Text)) <= Val(HBHV) And Val(Trim(txt_mncyteabs.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(txt_mncyteabs.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(txt_mncyteabs.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "MONOCYTE ABS"
                        UPDDHEMCBCDET_QR.Execute
                        
            End If

            
            
            If Len(Trim(txt_esnabs.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(txt_esnabs.Text)
                    HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Monocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Monocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_esnabs.Text)) <= Val(HBHV) And Val(Trim(txt_esnabs.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(txt_esnabs.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(txt_esnabs.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "EOSINOPHIL ABS"
                        UPDDHEMCBCDET_QR.Execute
                        
            End If

            
                        
            If Len(Trim(txt_bsphabs.Text)) > 0 Then
                
                
                    UPDDHEMCBCDET_QR(0) = Trim(txt_bsphabs.Text)
                    'HBAGE
                    BASOPHILSAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Basophils abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Basophils abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If
    
                    'RANGE = SELDHEMCBCRS!RANGE
                 
                    I = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    
                    If Val(Trim(txt_bsphabs.Text)) <= Val(HBHV) And Val(Trim(txt_bsphabs.Text)) >= Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "N"
                    ElseIf Val(Trim(txt_bsphabs.Text)) < Val(HBLV) Then
                        UPDDHEMCBCDET_QR(1) = "L"
                    ElseIf Val(Trim(txt_bsphabs.Text)) > Val(HBHV) Then
                        UPDDHEMCBCDET_QR(1) = "H"
                    End If
                        
                        UPDDHEMCBCDET_QR(2) = ""
                        UPDDHEMCBCDET_QR(3) = ""
                        UPDDHEMCBCDET_QR(4) = SELDHEMCBCRS("MACHINNA")
                        UPDDHEMCBCDET_QR(5) = SELDHEMCBCRS("UOM")
                        UPDDHEMCBCDET_QR(6) = SELDHEMCBCRS("RANGE")
                        UPDDHEMCBCDET_QR(7) = ""
                        UPDDHEMCBCDET_QR(8) = req_chglen(cmb_reqno.Text)
                        UPDDHEMCBCDET_QR(9) = "BASOPHIL ABS"
                        UPDDHEMCBCDET_QR.Execute
                        
            End If

End Sub


Public Sub LYMPHPCNTAGE()

If (txt_age.Text <= 0.02 And txt_age.Text >= "0.00") And Trim(txt_sex.Text) = "M" Then
    HAGE = 0.02
    HBSEX = "M"
End If

If (txt_age.Text <= 0.02 And txt_age.Text >= "0.00") And Trim(txt_sex.Text) = "F" Then
    HAGE = 0.02
    HBSEX = "F"
End If

If (txt_age.Text <= 0.1 And txt_age.Text >= 0.03) And Trim(txt_sex.Text) = "M" Then
    HAGE = 0.1
    HBSEX = "M"
End If

If (txt_age.Text <= 0.1 And txt_age.Text >= 0.03) And Trim(txt_sex.Text) = "F" Then
    HAGE = 0.1
    HBSEX = "F"
End If

If (txt_age.Text <= 0.2 And txt_age.Text > 0.1) And Trim(txt_sex.Text) = "M" Then
    HAGE = 0.2
    HBSEX = "M"
End If

If (txt_age.Text <= 0.2 And txt_age.Text > 0.1) And Trim(txt_sex.Text) = "F" Then
    HAGE = 0.2
    HBSEX = "F"
End If

If (txt_age.Text <= 0.6 And txt_age.Text > 0.2) And Trim(txt_sex.Text) = "M" Then
    HAGE = 0.6
    HBSEX = "M"
End If

If (txt_age.Text <= 0.6 And txt_age.Text > 0.2) And Trim(txt_sex.Text) = "F" Then
    HAGE = 0.6
    HBSEX = "F"
End If

If (txt_age.Text <= 6 And txt_age.Text > 0.6) And Trim(txt_sex.Text) = "M" Then
    HAGE = 6
    HBSEX = "M"
End If

If (txt_age.Text <= 6 And txt_age.Text > 0.6) And Trim(txt_sex.Text) = "F" Then
    HAGE = 6
    HBSEX = "F"
End If

If (txt_age.Text <= 12 And txt_age.Text > 6) And Trim(txt_sex.Text) = "M" Then
    HAGE = 12
    HBSEX = "M"
End If

If (txt_age.Text <= 12 And txt_age.Text > 6) And Trim(txt_sex.Text) = "F" Then
    HAGE = 12
    HBSEX = "F"
End If

If (txt_age.Text <= 106 And txt_age.Text > 12) And Trim(txt_sex.Text) = "M" Then
    HAGE = 106
    HBSEX = "M"
End If

If (txt_age.Text <= 106 And txt_age.Text > 12) And Trim(txt_sex.Text) = "F" Then
    HAGE = 106
    HBSEX = "F"
End If

End Sub



Private Sub TXT_TT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'SendKeys vbTab
KeyAscii = 0
End If
End Sub

Private Sub txt_ttctl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'SendKeys vbTab
KeyAscii = 0
End If
End Sub

