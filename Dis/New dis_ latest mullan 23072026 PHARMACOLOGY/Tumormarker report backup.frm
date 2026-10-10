VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form TumorMarkerRPTBackup 
   Caption         =   "Form2"
   ClientHeight    =   10110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15780
   LinkTopic       =   "Form2"
   ScaleHeight     =   10110
   ScaleWidth      =   15780
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame COMMAND 
      Height          =   735
      Left            =   30
      TabIndex        =   52
      Top             =   7935
      Width           =   11775
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Height          =   375
         Left            =   360
         MouseIcon       =   "Tumormarker report backup.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   61
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         Height          =   375
         Left            =   1620
         MouseIcon       =   "Tumormarker report backup.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   60
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   2850
         MouseIcon       =   "Tumormarker report backup.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4080
         MouseIcon       =   "Tumormarker report backup.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   58
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   5340
         MouseIcon       =   "Tumormarker report backup.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   57
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7995
         MouseIcon       =   "Tumormarker report backup.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   56
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   9240
         MouseIcon       =   "Tumormarker report backup.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         Height          =   375
         Left            =   6600
         MouseIcon       =   "Tumormarker report backup.frx":1546
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_addendum 
         Caption         =   "&Addendum"
         Height          =   375
         Left            =   10455
         TabIndex        =   53
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12360
      TabIndex        =   51
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12360
      TabIndex        =   50
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame common 
      Height          =   2565
      Left            =   30
      TabIndex        =   13
      Top             =   840
      Width           =   11790
      Begin VB.ComboBox cmb_dept 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         TabIndex        =   32
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txt_deptcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10200
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmb_reqno 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         TabIndex        =   30
         Top             =   615
         Width           =   2175
      End
      Begin VB.TextBox txt_pathno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   26
         Top             =   1455
         Width           =   735
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   21
         Top             =   1815
         Width           =   975
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
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   19
         Top             =   2160
         Width           =   3855
      End
      Begin VB.ComboBox cmb_modality 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         TabIndex        =   18
         Top             =   240
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   0
         TabIndex        =   15
         Top             =   720
         Width           =   1335
         Begin VB.OptionButton opt_staff 
            Caption         =   "Staff"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Opt_patient 
            Caption         =   "Patient"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1095
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
         TabIndex        =   14
         Top             =   2160
         Width           =   975
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   33
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483634
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
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         Caption         =   "Requisition Placed to"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   49
         Top             =   360
         Width           =   1830
      End
      Begin VB.Label lbl_deptcode 
         AutoSize        =   -1  'True
         Caption         =   "Dept. Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   48
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         Caption         =   "Case No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   47
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         Caption         =   "Requisition No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   46
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lbl_pathno 
         AutoSize        =   -1  'True
         Caption         =   "Lab No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9360
         TabIndex        =   45
         Top             =   600
         Width           =   690
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1680
         TabIndex        =   44
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   43
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5040
         TabIndex        =   42
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   41
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9720
         TabIndex        =   40
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         Caption         =   "Primary Doctor/Service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   1800
         Width           =   1995
      End
      Begin VB.Label lbl_ward 
         Caption         =   "Ward"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   38
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lbl_bed 
         AutoSize        =   -1  'True
         Caption         =   "Bed No."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6960
         TabIndex        =   37
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lbl_modality 
         AutoSize        =   -1  'True
         Caption         =   "Modality"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5400
         TabIndex        =   36
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl_Modalitycode 
         AutoSize        =   -1  'True
         Caption         =   "Modality Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   35
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         Caption         =   "Referring Doctor/service"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Width           =   2130
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   150
      TabIndex        =   2
      Top             =   7200
      Width           =   11685
      Begin VB.ComboBox cmb_enterby 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox cmb_finalby 
         Height          =   315
         Left            =   5055
         TabIndex        =   6
         Top             =   480
         Width           =   2010
      End
      Begin VB.TextBox txt_addendum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         TabIndex        =   5
         Top             =   150
         Width           =   3735
      End
      Begin VB.ComboBox cmb_addendumby 
         Height          =   315
         Left            =   8670
         TabIndex        =   4
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox TXT_COMMENTS 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label lbl_enter 
         AutoSize        =   -1  'True
         Caption         =   "Enter by"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl_finby 
         AutoSize        =   -1  'True
         Caption         =   "Finalized by"
         Height          =   195
         Left            =   3945
         TabIndex        =   11
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label LBL_COMM 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   120
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "Addendum"
         Height          =   255
         Left            =   6840
         TabIndex        =   9
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Addendum by"
         Height          =   240
         Left            =   7440
         TabIndex        =   8
         Top             =   510
         Width           =   1200
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11880
      Top             =   5520
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   11880
      Top             =   5880
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   11880
      Top             =   6360
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   11790
      ScaleHeight     =   4185
      ScaleWidth      =   3465
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox txtEInterface 
      Height          =   375
      Left            =   12030
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   11880
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   0
      TabIndex        =   62
      Top             =   3375
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   3030703
      TabCaption(0)   =   "Markers"
      TabPicture(0)   =   "Tumormarker report backup.frx":1850
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "REPORT"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Serum electrophoresis "
      TabPicture(1)   =   "Tumormarker report backup.frx":186C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FRMIMM"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FRMIMM 
         Caption         =   "Complete Serum Protein Electrophoresis (SPE) Profile "
         ForeColor       =   &H000000C0&
         Height          =   3375
         Left            =   120
         TabIndex        =   67
         Top             =   360
         Width           =   11655
         Begin VB.Frame Frame6 
            Caption         =   "Serum Light Chains"
            ForeColor       =   &H000000C0&
            Height          =   2655
            Left            =   3480
            TabIndex        =   96
            Top             =   240
            Width           =   3375
            Begin VB.TextBox txt_lambda 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               TabIndex        =   100
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txt_kappa 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               TabIndex        =   99
               Top             =   360
               Width           =   975
            End
            Begin VB.TextBox txt_kappafree 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               TabIndex        =   98
               Top             =   1560
               Width           =   975
            End
            Begin VB.TextBox txt_lambdafree 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               TabIndex        =   97
               Top             =   2160
               Width           =   975
            End
            Begin VB.Label Label9 
               Caption         =   "269 - 638"
               Height          =   255
               Left            =   2280
               TabIndex        =   108
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label8 
               Caption         =   "574 - 1276"
               Height          =   255
               Left            =   2280
               TabIndex        =   107
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label6 
               Caption         =   "Lambda Total mg/dl"
               Height          =   435
               Left            =   120
               TabIndex        =   106
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "Kappa Total mg/dl"
               Height          =   435
               Left            =   120
               TabIndex        =   105
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label10 
               Caption         =   "Kappa Free mg/l"
               Height          =   435
               Left            =   120
               TabIndex        =   104
               Top             =   1440
               Width           =   1050
            End
            Begin VB.Label Label11 
               Caption         =   "Lambda Free mg/l"
               Height          =   435
               Left            =   120
               TabIndex        =   103
               Top             =   1920
               Width           =   1170
            End
            Begin VB.Label Label15 
               Caption         =   "3.3 - 19.4"
               Height          =   255
               Left            =   2280
               TabIndex        =   102
               Top             =   1560
               Width           =   855
            End
            Begin VB.Label Label16 
               Caption         =   "5.71 - 26.3"
               Height          =   255
               Left            =   2280
               TabIndex        =   101
               Top             =   2160
               Width           =   975
            End
         End
         Begin VB.ComboBox cmb_polyclonal 
            Height          =   315
            Left            =   7920
            TabIndex        =   95
            Top             =   2880
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txt_monoclonal 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   94
            Top             =   3000
            Width           =   9975
         End
         Begin VB.ComboBox cmb_urineimm 
            Height          =   315
            Left            =   10560
            TabIndex        =   93
            Top             =   2040
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ComboBox cmb_ka 
            Height          =   315
            Left            =   10320
            TabIndex        =   92
            Top             =   2400
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame Frame2 
            Caption         =   "Serum Immunoglobulins"
            ForeColor       =   &H000000C0&
            Height          =   2655
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   3375
            Begin VB.TextBox TXT_IG 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               TabIndex        =   85
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox TXT_TM 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   195
               Left            =   1320
               TabIndex        =   84
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox TXT_IGG1 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               TabIndex        =   83
               Top             =   2040
               Width           =   975
            End
            Begin VB.TextBox txt_igm 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               TabIndex        =   82
               Top             =   1020
               Width           =   975
            End
            Begin VB.TextBox txt_igg 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1200
               TabIndex        =   81
               Top             =   1560
               Width           =   1095
            End
            Begin VB.TextBox txt_iga 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   1200
               TabIndex        =   80
               Top             =   480
               Width           =   975
            End
            Begin VB.Label lbl_iga 
               AutoSize        =   -1  'True
               Caption         =   "IgA  mg/dl "
               Height          =   195
               Left            =   120
               TabIndex        =   91
               Top             =   480
               Width           =   960
            End
            Begin VB.Label lbl_igm 
               AutoSize        =   -1  'True
               Caption         =   "IgM  mg/dl "
               Height          =   195
               Left            =   120
               TabIndex        =   90
               Top             =   1020
               Width           =   990
            End
            Begin VB.Label lbl_igg 
               AutoSize        =   -1  'True
               Caption         =   "IgG  mg/dl "
               Height          =   195
               Left            =   120
               TabIndex        =   89
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label LBL_RANGE 
               Caption         =   "82 - 453"
               Height          =   255
               Left            =   2280
               TabIndex        =   88
               Top             =   480
               Width           =   855
            End
            Begin VB.Label LBL_IGMR 
               Caption         =   "46 - 304"
               Height          =   255
               Left            =   2280
               TabIndex        =   87
               Top             =   1020
               Width           =   735
            End
            Begin VB.Label LBL_IGGR 
               Caption         =   "751 - 1560"
               Height          =   255
               Left            =   2280
               TabIndex        =   86
               Top             =   1560
               Width           =   975
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Serum Protein Electrophoresis"
            ForeColor       =   &H000000C0&
            Height          =   2655
            Left            =   6840
            TabIndex        =   68
            Top             =   240
            Width           =   4695
            Begin VB.ComboBox cmb_elepattern 
               Height          =   315
               ItemData        =   "Tumormarker report backup.frx":1888
               Left            =   2280
               List            =   "Tumormarker report backup.frx":188A
               TabIndex        =   73
               Top             =   360
               Width           =   1815
            End
            Begin VB.ComboBox cmb_locationnband 
               Height          =   315
               Left            =   2280
               TabIndex        =   72
               Top             =   780
               Width           =   1815
            End
            Begin VB.ComboBox cmb_monband 
               Height          =   315
               ItemData        =   "Tumormarker report backup.frx":188C
               Left            =   2280
               List            =   "Tumormarker report backup.frx":188E
               TabIndex        =   71
               Top             =   1215
               Width           =   2055
            End
            Begin VB.TextBox txt_connband 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2280
               TabIndex        =   70
               Top             =   1635
               Width           =   2295
            End
            Begin VB.TextBox txt_remmbond 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2280
               TabIndex        =   69
               Top             =   2160
               Width           =   2295
            End
            Begin VB.Label lbl_elepattern 
               AutoSize        =   -1  'True
               Caption         =   "Electrophoretic pattern"
               Height          =   195
               Left            =   120
               TabIndex        =   78
               Top             =   360
               Width           =   1965
            End
            Begin VB.Label lbl_monband 
               AutoSize        =   -1  'True
               Caption         =   "Location of M band"
               Height          =   195
               Left            =   120
               TabIndex        =   77
               Top             =   810
               Width           =   1665
            End
            Begin VB.Label lbl_monoband1 
               AutoSize        =   -1  'True
               Caption         =   "Monoclonal Band"
               Height          =   195
               Left            =   120
               TabIndex        =   76
               Top             =   1260
               Width           =   1485
            End
            Begin VB.Label lbl_conmband 
               AutoSize        =   -1  'True
               Caption         =   "Concentration of M band"
               Height          =   195
               Left            =   120
               TabIndex        =   75
               Top             =   1710
               Width           =   2115
            End
            Begin VB.Label lbl_remmbond 
               AutoSize        =   -1  'True
               Caption         =   "Remarks "
               Height          =   195
               Left            =   120
               TabIndex        =   74
               Top             =   2160
               Width           =   810
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Urine ( 24 Hours )"
            ForeColor       =   &H000000C0&
            Height          =   3375
            Left            =   0
            TabIndex        =   109
            Top             =   0
            Visible         =   0   'False
            Width           =   10575
            Begin VB.ComboBox CMB_BJPBJP 
               Height          =   315
               ItemData        =   "Tumormarker report backup.frx":1890
               Left            =   3360
               List            =   "Tumormarker report backup.frx":189A
               TabIndex        =   114
               Top             =   480
               Width           =   2295
            End
            Begin VB.TextBox txt_urrate 
               Height          =   285
               Left            =   3360
               TabIndex        =   113
               Top             =   960
               Width           =   4215
            End
            Begin VB.TextBox txt_urkappa 
               Height          =   285
               Left            =   3360
               TabIndex        =   112
               Top             =   1680
               Width           =   2055
            End
            Begin VB.TextBox txt_urlambda 
               Height          =   285
               Left            =   3360
               TabIndex        =   111
               Top             =   2040
               Width           =   2055
            End
            Begin VB.ComboBox cmb_urrate 
               Height          =   315
               Left            =   3360
               TabIndex        =   110
               Text            =   "Combo1"
               Top             =   960
               Width           =   4215
            End
            Begin VB.Label lbl_benjone 
               AutoSize        =   -1  'True
               Caption         =   "Bence Jones Protein BJP"
               Height          =   195
               Left            =   480
               TabIndex        =   120
               Top             =   480
               Width           =   2160
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Urine Excretion Rate"
               Height          =   195
               Left            =   480
               TabIndex        =   119
               Top             =   960
               Width           =   1785
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Urine Free Light Chains Kappa"
               Height          =   195
               Left            =   480
               TabIndex        =   118
               Top             =   1680
               Width           =   2610
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Urine Free Light Chains Lambda"
               Height          =   195
               Left            =   480
               TabIndex        =   117
               Top             =   2040
               Width           =   2730
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "5.71 - 26.3   mg/l"
               Height          =   195
               Left            =   5640
               TabIndex        =   116
               Top             =   1680
               Width           =   1500
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "0.81 - 10.1   mg/l"
               Height          =   195
               Left            =   5640
               TabIndex        =   115
               Top             =   2040
               Width           =   1500
            End
         End
         Begin VB.Label lbl_urineimm 
            AutoSize        =   -1  'True
            Caption         =   "Urine Immunoelectropho."
            Height          =   195
            Left            =   9120
            TabIndex        =   124
            Top             =   2160
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label lbl_typmono 
            AutoSize        =   -1  'True
            Caption         =   "Immunofixation"
            Height          =   195
            Left            =   120
            TabIndex        =   123
            Top             =   3000
            Width           =   1275
         End
         Begin VB.Label lbl_kp 
            AutoSize        =   -1  'True
            Caption         =   "Kappa / Lamda"
            Height          =   195
            Left            =   9000
            TabIndex        =   122
            Top             =   2520
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label lbl_polyclonal 
            AutoSize        =   -1  'True
            Caption         =   "Polyclonal"
            Height          =   195
            Left            =   6960
            TabIndex        =   121
            Top             =   3000
            Visible         =   0   'False
            Width           =   885
         End
      End
      Begin VB.Frame REPORT 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   63
         Top             =   360
         Width           =   11535
         Begin VB.TextBox txt_observed 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            Height          =   285
            Left            =   3480
            TabIndex        =   65
            Top             =   480
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.ComboBox cmb_machine 
            Height          =   315
            ItemData        =   "Tumormarker report backup.frx":18B2
            Left            =   10320
            List            =   "Tumormarker report backup.frx":18CE
            TabIndex        =   64
            Top             =   480
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   2895
            Left            =   120
            TabIndex        =   66
            Top             =   120
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   5106
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            ForeColorFixed  =   192
            AllowUserResizing=   3
            FormatString    =   "S.No.|^Procedure Name                   |^Observed Value|^Normal Range of the Procedure |^ Panic Value|^Source|^Preval|^MachineNo."
         End
      End
   End
   Begin Crystal.CrystalReport CR_SubTM 
      Left            =   30
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblicu 
      Caption         =   "ICU Patient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   120
      TabIndex        =   133
      Top             =   495
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lbl_rwd 
      Caption         =   "RECOVERY WARD"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   150
      TabIndex        =   132
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   11760
      TabIndex        =   131
      Top             =   0
      Width           =   420
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   11760
      TabIndex        =   130
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lbl_biochemlab 
      AutoSize        =   -1  'True
      Caption         =   "REPORT - TUMOUR MARKER SERVICES "
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
      Left            =   3510
      TabIndex        =   129
      Top             =   480
      Width           =   5700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Location"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6630
      TabIndex        =   128
      Top             =   3120
      Width           =   750
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
      Left            =   1680
      TabIndex        =   127
      Top             =   0
      Width           =   9255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Immunofixation"
      Height          =   195
      Left            =   390
      TabIndex        =   126
      Top             =   6000
      Width           =   1275
   End
   Begin VB.Label Label19 
      Caption         =   "05/11/21"
      Height          =   255
      Left            =   600
      TabIndex        =   125
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "TumorMarkerRPTBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ADDED BY MILIND DT. 03/03/06
Dim userid_qr As rdoQuery
Dim userid_rs As rdoResultset
Dim userdesc_qr As rdoQuery
Dim userdesc_rs As rdoResultset
Dim COMMENTLIST_qr As rdoQuery
Dim COMMENTLIST_rs As rdoResultset
'upto here
'txt_kappa txt_lambda
'ADDED BY SHRIKANT FOR BINTERFACE
Dim rs_binterface As rdoResultset
'******** Shraddha 11/02/09 *********
Dim chk_var As Boolean

'***** Shraddha 05/01/11 *****
Dim SELDBCALKPHSQR As rdoQuery
Dim SELDBCALKPHSRS As rdoResultset
'************************************
'***** ShWETA 11/05/12 *****
Dim VWREQNSTAT As String
'************************************
Dim ADDENDUMFALG As Boolean
Dim IMFLAG As Boolean
Dim j As Integer
Dim K As Integer
Dim l As Integer
Dim tmflag As Boolean
Dim str As String
'Dim FSFLAG As Boolean
Dim sel_deptqr As rdoQuery
'Dim UPDSPREPORT_QR As rdoQuery
Dim sel_deptrs As rdoResultset
Dim SEL_DEPTRSFLAG As Boolean
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
Dim SELPATHDREQUISITN_QR As rdoQuery
Dim selpathdrequisitn_rs As rdoResultset
Dim SELPATHDREQUISITNFALG1 As Boolean


'Dim SELEMPLIB_QR As rdoQuery
'Dim SELEMPLIB_RS As rdoResultset

Dim i As Integer
'Dim insdfsrep_qr As rdoQuery
'Dim insdgrossdesc_qr As rdoQuery
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

Dim seldissecentr1_qr As rdoQuery
Dim seldissecentr1_rs As rdoResultset
Dim seldissecentrflag1 As Boolean

Dim seldissecfin1_qr As rdoQuery
Dim seldissecfin1_rs As rdoResultset
Dim seldissecfinflag1 As Boolean

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
Dim INSDIMBIOCHEQR As rdoQuery
Dim UPDDIMBIOCHEQR As rdoQuery
Dim SELDIMBIOCHEQR As rdoQuery
Dim SELDIMBIOCHERS As rdoResultset
Dim SELDIMBIOCHEFLAG As Boolean

Dim SELDIMBIOCHEPQR As rdoQuery
Dim SELDIMBIOCHEPRS As rdoResultset
Dim SELDIMBIOCHEPFLAG As Boolean


Dim INSDMBCCOMMQR As rdoQuery
Dim UPDDMBCCOMMQR As rdoQuery
Dim SELDMBCCOMMQR As rdoQuery
Dim SELDMBCCOMMRS As rdoResultset
Dim SELDMBCCOMMFLAG As Boolean
Dim updrrequisitn1_qr As rdoQuery

''*******ICU**********
Dim VAR_REQNNO As String
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
Dim REPWTRIRS As rdoResultset
Dim CURUSRDET_QR As rdoQuery
Dim CURUSRDET_RS As rdoResultset

'RTIMELOG
Dim UPDRTIMECADD_QR As rdoQuery
Dim UPDRTIME_QR As rdoQuery
Dim UPDRTIMEC_QR As rdoQuery
Dim UPDRTIMEP_QR As rdoQuery
Dim UPDRTIMEM_QR As rdoQuery

'addendum
Dim insertaddendum_qr As rdoQuery
Dim selectaddendum_qr As rdoQuery
Dim selectaddendum_rs As rdoResultset
Dim Addendum_Flag As Boolean

'INTERFACE
Dim SELINTEREXPCDQR As rdoQuery
Dim SELINTEREXPCDRS As rdoResultset
Dim SELINTEREXPCDFLAG As Boolean

Dim SELDTEMPLETQR1 As rdoQuery
Dim SELDTEMPLETRS1 As rdoResultset
Dim SELDTEMPLETFLAG1 As Boolean

Dim INTERFACEREAD_QR As rdoQuery
Dim INTERFACEREAD_RS As rdoResultset

Dim UPDINTERFACE_QR As rdoQuery


Dim SELFINALBY_QR As rdoQuery
Dim SELFINALBY_RS As rdoResultset
Dim REPWTRN_QR As rdoQuery


'''KANCHAN
Dim SELINTER_FLAG As Boolean
Dim SELINTER_QR As rdoQuery
Dim SELINTER_RS As rdoResultset
Dim PC As String
Dim selMA101QR As rdoQuery
Dim SELMA101RS As rdoResultset
Dim TEST_DET As String
Dim INSDIMBIOCHE1 As rdoQuery
Dim UPDDIMBIOCHE1 As rdoQuery

Dim CLPTIME As String

'********** Shraddha 23/01/09 *********
Dim SEL_BIOCHE1_QR As rdoQuery
Dim SEL_BIOCHE1_RS As rdoResultset

Dim sel_dinterface_qr As rdoQuery
Dim sel_dinterface_rs As rdoResultset

Dim upd_dinterface_qr As rdoQuery

Dim reqno_dinter As String
'**************************************
Dim SEL_DINT_QR As rdoQuery 'Shraddha 07/04/10
Dim SEL_DINT_RS As rdoResultset 'Shraddha 07/04/10

Dim UPD_DInterface As rdoQuery 'Shraddha 28/07/10
Dim RS As rdoResultset
Dim STRSQL As String


''******SECURITY***********
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
    'cmd_fullreport.Enabled = False
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    If USRNAME_RS("MODIFY") = "YES" Then
    'shweta 11.05.2012 as told by kalpana mam
        cmd_modify.Enabled = True
        'cmd_modify.Enabled = True
    Else
        cmd_modify.Enabled = False
    End If
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = True
   ' cmd_commit.Enabled = False
   If USRNAME_RS("COMMIT") = "YES" Then
        cmd_commit.Enabled = True
    Else
        cmd_commit.Enabled = False
    End If
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
    cmd_addendum.Visible = False
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


'If cap = "NEW" Then
'    cmd_save.Enabled = True
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    CMD_NEW.Enabled = False
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = False
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = False
'    'cmd_fullreport.Enabled = False
'ElseIf cap = "VIEW" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = True
'    CMD_NEW.Enabled = False
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = False
'    'cmd_fullreport.Enabled = True
'ElseIf cap = "SAVE" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    CMD_NEW.Enabled = True
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = True
'    'cmd_fullreport.Enabled = False
'ElseIf cap = "INI" Then
'    cmd_cancel.Enabled = False
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    CMD_NEW.Enabled = True
'    cmd_save.Enabled = False
'    cmd_print.Enabled = False
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = True
'   ' cmd_fullreport.Enabled = False
'ElseIf cap = "COMMIT" Then
'    cmd_save.Enabled = False
'    cmd_cancel.Enabled = False
'    cmd_back.Enabled = True
'    CMD_NEW.Enabled = True
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = True
'    'cmd_fullreport.Enabled = False
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

Public Sub FLDENABL(str1 As String)
If str1 = "ENA" Then
    cmb_modality.Enabled = True
    cmb_reqno.Enabled = True
    cmb_dept.Enabled = True
    MaskEdBox1.Enabled = True
    MSFlexGrid1.Enabled = True
    cmb_elepattern.Enabled = True
     cmb_monband.Enabled = True
     cmb_locationnband.Enabled = True
     'cmb_connband.Enabled = True
     txt_iga.Enabled = True
     txt_igm.Enabled = True
     txt_igg.Enabled = True
     txt_kappa.Enabled = True
     txt_lambda.Enabled = True
     txt_remmbond.Enabled = True
     
     cmb_urineimm.Enabled = True
     CMB_BJPBJP.Enabled = True
     cmb_ka.Enabled = True
     txt_monoclonal.Enabled = True
      
     txt_comments.Enabled = True
     cmb_polyclonal.Enabled = True
                  
    
End If
If str1 = "DIS" Then
    cmb_modality.Enabled = False
    cmb_reqno.Enabled = False
    cmb_dept.Enabled = False
    MaskEdBox1.Enabled = False
    
    MSFlexGrid1.Enabled = False
    cmb_monband.Enabled = False
     cmb_locationnband.Enabled = False
     cmb_elepattern.Enabled = False
     txt_iga.Enabled = False
     txt_igm.Enabled = False
     txt_igg.Enabled = False
     txt_kappa.Enabled = False
     txt_lambda.Enabled = False
     txt_remmbond.Enabled = False
     cmb_urineimm.Enabled = False
     CMB_BJPBJP.Enabled = False
     cmb_ka.Enabled = False
     txt_monoclonal.Enabled = False
     txt_comments.Enabled = False
     cmb_polyclonal.Enabled = False
End If
If str1 = "VIEW" Then
    cmb_modality.Enabled = True
    cmb_reqno.Enabled = True
    cmb_dept.Enabled = True
    MaskEdBox1.Enabled = True
    MSFlexGrid1.Enabled = False
     cmb_monband.Enabled = False
     cmb_locationnband.Enabled = False
     cmb_elepattern.Enabled = False
     txt_iga.Enabled = False
     txt_igm.Enabled = False
     txt_igg.Enabled = False
     txt_kappa.Enabled = False
     txt_lambda.Enabled = False
     txt_remmbond.Enabled = False
     cmb_urineimm.Enabled = False
     CMB_BJPBJP.Enabled = False
     cmb_ka.Enabled = False
     txt_monoclonal.Enabled = False
     txt_comments.Enabled = False
     cmb_polyclonal.Enabled = False

End If
If str1 = "MODIFY" Then
    cmb_modality.Enabled = False
    cmb_reqno.Enabled = False
    cmb_dept.Enabled = False
    MaskEdBox1.Enabled = False
    'TXT_CASENO.Enabled = False
    'txt_Location.Enabled = True
    'txt_amtchg.Enabled = True
    'txt_recpt.Enabled = True
    MSFlexGrid1.Enabled = True
     cmb_monband.Enabled = True
     cmb_locationnband.Enabled = True
     cmb_elepattern.Enabled = True
     txt_iga.Enabled = True
     txt_igm.Enabled = True
     txt_igg.Enabled = True
     txt_kappa.Enabled = True
     txt_lambda.Enabled = True
     txt_remmbond.Enabled = True
     cmb_urineimm.Enabled = True
     CMB_BJPBJP.Enabled = True
     cmb_ka.Enabled = True
     txt_monoclonal.Enabled = True
     txt_comments.Enabled = True
    cmb_polyclonal.Enabled = True
End If
End Sub





Private Sub CMB_BJPBJP_Click()
If CMB_BJPBJP.Text = "Positive" Then
cmb_urrate.SetFocus
'SendKeys "^{F4}"
Else
cmb_urrate.Text = ""
End If
End Sub

Private Sub cmb_machine_Click()
Dim num As Integer
Dim sc As String

cmb_machine.Visible = False
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = Trim(cmb_machine.Text)
num = InStr(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), "$")
sc = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), num + 1)

     If SELDTEMPLETFLAG1 = False Then
         SELDTEMPLETQR1(0) = sc
         SELDTEMPLETQR1(1) = Trim(cmb_machine.Text)
         Set SELDTEMPLETRS1 = SELDTEMPLETQR1.OpenResultset(1, 3)
         SELDTEMPLETFLAG1 = True
     Else
         SELDTEMPLETQR1(0) = sc
         SELDTEMPLETQR1(1) = Trim(cmb_machine.Text)
         SELDTEMPLETRS1.MoreResults
         SELDTEMPLETRS1.Requery
     End If
     
     If SELDTEMPLETRS1.RowCount > 0 Then
        RNAGE1 = Trim(SELDTEMPLETRS1!LOWVALUE) & "-" & Trim(SELDTEMPLETRS1!HIGHVALUE)
        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = RNAGE1
     Else
        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = ""
     End If
    '***** Shraddha 06/01/11 *****
    If Trim(sc) = "F821" Then
        '**** Age *****
        If txt_age.Text >= 1 And txt_age.Text <= 60 Then
            RNAGE = 60
        End If
        If (txt_age.Text >= 61 And txt_age.Text <= 100) Then
            RNAGE = 100
        End If
        '*************
        SELDBCALKPHSQR(0) = "F821"
        SELDBCALKPHSQR(1) = RNAGE
        SELDBCALKPHSQR(2) = Trim(cmb_machine.Text)
        Set SELDBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
        If Not SELDBCALKPHSRS.EOF Then
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = Trim(SELDBCALKPHSRS!range)
        End If
    End If
    '*******************************
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
        If Trim(cmb_modality.List(ind_cbodept)) = "TUMOUR MARKER$TM" Then
            cmb_modality.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
    End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub



Private Sub cmb_polyclonal_GotFocus()
cmb_polyclonal.clear
cmb_polyclonal.AddItem "Present"
cmb_polyclonal.AddItem "Absent"

End Sub

'''''DONE BY KANCHAN TO CALCULATE LAPS TIME FOR TEST F833,F834,F35,F836
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
Public Sub SELDINTERFACE()

    If SELINTER_FLAG = False Then
        SELINTER_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
        Set SELINTER_RS = SELINTER_QR.OpenResultset(1, 3)
        SELINTER_FLAG = True
    Else
        SELINTER_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
        SELINTER_RS.MoreResults
        SELINTER_RS.Requery
    End If
    If SELINTER_RS.RowCount > 0 Then
        REPORT.Enabled = True
        i = 1
        MSFlexGrid1.Rows = 1

        Do Until SELINTER_RS.EOF
            PC = SELINTER_RS!TESTCODE
        
            'THIS IS FOR GETTING TEST DESCRIPTION FOR TEST CODE
            SELMA101FLAG = False
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
            
            
            '***** Shraddha 07/04/10 *****
            'SELECT * FROM interface.dinterface where reqno=? and TESTCODE=?
            SEL_DINT_QR(0) = Trim(req_chglen(cmb_reqno.Text))
            SEL_DINT_QR(1) = PC
            Set SEL_DINT_RS = SEL_DINT_QR.OpenResultset(1, 3)
            'SELDTEMPLETQR(1) = Trim(SEL_DINT_RS("MACHINENO"))
            '*****************************
            
            
            'THIS IS FOR GETTING HIGHVALUE AND LOWVALUE OF TESTCODE
            'If PC <> "F053" And PC <> "FO54" And PC <> "F977" Then
            'Added by Shraddha on 13/04/10
'            str = "SELECT * FROM DISLIB.DTEMPLET WHERE PRO_CODE=?"
'            Set SELDTEMPLETQR = rdb.CreateQuery("", str)
            
            'If SELDTEMPLETFLAG = False Then
'                SELDTEMPLETQR(0) = PC
'                Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
            '    SELDTEMPLETFLAG = True
'            Else
'                SELDTEMPLETQR(0) = PC
'                SELDTEMPLETRS.MoreResults
'                SELDTEMPLETRS.Requery
'            End If
            
'            If SELDTEMPLETFLAG = False Then
                SELDTEMPLETQR(0) = PC
                'SELDTEMPLETQR(1) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)) '07/04/10
                SELDTEMPLETQR(1) = Trim(SEL_DINT_RS("MACHINENO"))
                
                If Trim(SEL_DINT_RS("MACHINENO")) = "" Then
                    'MsgBox "Machine name has not been inserted in Interface"
                
                '***** Add code to update relevent rec with dummy m/c name *****
                'str = "UPDATE interface.dinterface SET MACHINENO=? WHERE REQNO=? and TESTCODE=?"
                UPD_DInterface(0) = "ARCH"
                UPD_DInterface(1) = Trim(req_chglen(cmb_reqno.Text))
                UPD_DInterface(2) = PC
                UPD_DInterface.Execute

                SEL_DINT_QR(0) = Trim(req_chglen(cmb_reqno.Text))
                SEL_DINT_QR(1) = PC
                Set SEL_DINT_RS = SEL_DINT_QR.OpenResultset(1, 3)

                SELDTEMPLETQR(0) = PC
                SELDTEMPLETQR(1) = Trim(SEL_DINT_RS("MACHINENO"))
                '***************************************************************
                
                End If
                
                Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                
                
                
                
'                SELDTEMPLETFLAG = True
'            Else
'                SELDTEMPLETQR(0) = PC
'                'SELDTEMPLETQR(1) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
'                SELDTEMPLETQR(1) = Trim(SEL_DINT_RS("MACHINENO"))
'                SELDTEMPLETRS.MoreResults
'                SELDTEMPLETRS.Requery
'            End If
                
            If SELDTEMPLETRS.RowCount > 0 Then ''IF 1
                RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
                'PC = Trim(DMULPROCE_RS("pro_code"))
                
                If PC <> "F816" And PC <> "F828" And PC <> "F977" And PC <> "F046" And PC <> "F047" And PC <> "F048" And PC <> "F050" And PC <> "F051" And PC <> "F041" And PC <> "F042" And PC <> "F053" And PC <> "F054" Then '' IF2
     
                    If SELINTEREXPCDFLAG = False Then   'CHECKING FOR EXCETION VALUES
                        SELINTEREXPCDQR(0) = PC
                        Set SELINTEREXPCDRS = SELINTEREXPCDQR.OpenResultset(1, 3)
                        SELINTEREXPCDFLAG = True
                    Else
                        SELINTEREXPCDQR(0) = PC
                        SELINTEREXPCDRS.MoreResults
                        SELINTEREXPCDRS.Requery
                    End If
                                 
                    If SELINTEREXPCDRS.RowCount > 0 Then
                       'MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS(1)) & "$" & PC '& vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                    Else
                        'THIS IS FOR PICKUP TESTRESULT FROM INTERFACE/DINTERFACE
                        INTERFACEREAD_QR(1) = PC
                        INTERFACEREAD_QR(0) = VAR_REQNNO
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
                            'MSFlexGrid1.AddItem i & vbTab & Trim(selMA101RS(1)) & "$" & PC & vbTab & INTFAC & vbTab & RNAGE1 & vbTab & vbTab & INTMODE & vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL"))) & vbTab & MACHINENO
                        Else
                            INTFAC = ""
                            INTMODE = ""
                            'MSFlexGrid1.AddItem i & vbTab & Trim(selMA101RS(1)) & "$" & PC '& vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                        End If
                        
                    End If
                    'Next statement commented by Shraddha on 11/02/09
                    'MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS(1)) & "$" & PC & vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
                    'SELINTER_QR(0) = VAR_REQNNO
                    'Set SELINTER_RS = SELINTER_QR.OpenResultset(1, 3)
                    
                    '***** Shraddha 06/01/11 *****
                    Dim range As String
                    If PC = "F821" Then
                        '**** Age *****
                        If txt_age.Text >= 1 And txt_age.Text <= 60 Then
                            RNAGE = 60
                        End If
                        If (txt_age.Text >= 61 And txt_age.Text <= 100) Then
                            RNAGE = 100
                        End If
                        '*************
                        SELDBCALKPHSQR(0) = "F821"
                        SELDBCALKPHSQR(1) = RNAGE 'Trim(txt_age.Text)
                        SELDBCALKPHSQR(2) = SELINTER_RS("MACHINENO") 'Trim(DMPROCEDUR_RS("REMARKS"))
                        Set SELDBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
                        If Not SELDBCALKPHSRS.EOF Then
                            range = Trim(SELDBCALKPHSRS!range)
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = range
                        End If
                        
                        MsgBox "machine name " & SELINTER_RS("MACHINENO")
                    End If
                    '*****************************
                    'Next if condition added by Shraddha Pandit on 11/02/09
                    If Trim(SELINTER_RS("TESTRESULT")) <> "" Then
                    
                        'Machine no added by Shraddha Pandit on 23/04/09 in else part too
                        MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS(1)) & "$" & PC & vbTab & SELINTER_RS("TESTRESULT") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & vbTab & vbTab & vbTab & SELINTER_RS("MACHINENO")
                    Else
                        MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS(1)) & "$" & PC & vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & vbTab & vbTab & vbTab & SELINTER_RS("MACHINENO")
                    End If
                    i = i + 1
                                
                    
                Else
                    IMFLAG = True
                End If ''IF 2
                            
                            
            Else
                
                '***** Added by Shraddha on 20/05/10 *****
'                Dim str1
'                str1 = "SELECT * FROM DISLIB.DTEMPLET WHERE PRO_CODE=?"
'                Set SELDTEMPLETQR = rdb.CreateQuery("", str1)
'
'                SELDTEMPLETQR(0) = PC
'                Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                
                'If Trim(SEL_DINT_RS("MACHINENO")) <> "" Then
                If Not SELDTEMPLETRS.EOF Then
                    SELDTEMPLETQR(0) = PC
                    'SELDTEMPLETQR(1) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)) '07/04/10
                    SELDTEMPLETQR(1) = Trim(SEL_DINT_RS("MACHINENO"))
                    
                    If Trim(SEL_DINT_RS("MACHINENO")) = "" Then
                        MsgBox "Machine name has not been inserted in Interface"
                    End If
                    
                    Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                   
                    MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101RS(1)) & "$" & PC & vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & vbTab & vbTab & vbTab & SELINTER_RS("MACHINENO")
               End If
                'str = ""
                '*****************************************
            
                IMFLAG = True
                FRMIMM.Enabled = True
                cmb_elepattern.clear
                cmb_monband.clear
                cmb_locationnband.clear
                cmb_urineimm.clear
                CMB_BJPBJP.clear
                cmb_ka.clear
                TMADD
                             
            End If ''IF1
                    
            SELINTER_RS.MoveNext
            
        Loop
            
            
        SELINTER_RS.MoveFirst
        While Not SELINTER_RS.EOF
            If PC <> "F816" And PC <> "F828" And PC <> "F977" And PC <> "F046" And PC <> "F047" And PC <> "F048" And PC <> "F050" And PC <> "F051" And PC <> "F041" And PC <> "F042" And PC <> "F053" And PC <> "F054" Then
                txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                SSTab1.TabVisible(1) = False
            End If
            SELINTER_RS.MoveNext
        Wend
        
        If IMFLAG = True Then
            SSTab1.TabVisible(1) = True
        Else
            SSTab1.TabVisible(1) = False
        End If
               
    End If
                           
End Sub

Public Sub dinter()
'***** Shraddha 08/12/09 *****
sel_dinterface_qr(0) = req_chglen(cmb_reqno.Text)
Set sel_dinterface_rs = sel_dinterface_qr.OpenResultset(1, 3)
If Not sel_dinterface_rs.EOF Then
    sel_dinterface_rs.MoveFirst
    reqno_dinter = sel_dinterface_rs("reqno")
End If

While Not sel_dinterface_rs.EOF
    FRMIMM.Visible = True
    
    If sel_dinterface_rs("TESTCODE") = "F046" Then
    
        If Trim(sel_dinterface_rs("TESTRESULT")) <> "" Then
        If InStr(1, Trim(sel_dinterface_rs("TESTRESULT")), ".", vbBinaryCompare) = 0 Then
            txt_iga.Text = Trim(sel_dinterface_rs("TESTRESULT")) & ".00"
        Else
            txt_iga.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
        Else
        txt_iga.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
    ElseIf sel_dinterface_rs("TESTCODE") = "F047" Then
        If Trim(sel_dinterface_rs("TESTRESULT")) <> "" Then
        If InStr(1, Trim(sel_dinterface_rs("TESTRESULT")), ".", vbBinaryCompare) = 0 Then
            txt_igm.Text = Trim(sel_dinterface_rs("TESTRESULT")) & ".00"
        Else
            txt_igm.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
        Else
            txt_igm.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
        
    ElseIf sel_dinterface_rs("TESTCODE") = "F048" Then
        'txt_igg.Text = sel_dinterface_rs("TESTRESULT")
        If Trim(sel_dinterface_rs("TESTRESULT")) <> "" Then
        If InStr(1, Trim(sel_dinterface_rs("TESTRESULT")), ".", vbBinaryCompare) = 0 Then
            txt_igg.Text = Trim(sel_dinterface_rs("TESTRESULT")) & ".00"
        Else
            txt_igg.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
        Else
            txt_igg.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
        
    ElseIf sel_dinterface_rs("TESTCODE") = "F050" Then
        'txt_kappa.Text = sel_dinterface_rs("TESTRESULT")
        If Trim(sel_dinterface_rs("TESTRESULT")) <> "" Then
        If InStr(1, Trim(sel_dinterface_rs("TESTRESULT")), ".", vbBinaryCompare) = 0 Then
            txt_kappa.Text = Trim(sel_dinterface_rs("TESTRESULT")) & ".00"
        Else
            txt_kappa.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
        Else
            txt_kappa.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
        
    ElseIf sel_dinterface_rs("TESTCODE") = "F051" Then
        'txt_lambda.Text = sel_dinterface_rs("TESTRESULT")
        If Trim(sel_dinterface_rs("TESTRESULT")) <> "" Then
        If InStr(1, Trim(sel_dinterface_rs("TESTRESULT")), ".", vbBinaryCompare) = 0 Then
            txt_lambda.Text = Trim(sel_dinterface_rs("TESTRESULT")) & ".00"
        Else
            txt_lambda.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
        Else
            txt_lambda.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
        
    ElseIf sel_dinterface_rs("TESTCODE") = "F041" Then
        'txt_kappafree.Text = sel_dinterface_rs("TESTRESULT")
        If Trim(sel_dinterface_rs("TESTRESULT")) <> "" Then
        If InStr(1, Trim(sel_dinterface_rs("TESTRESULT")), ".", vbBinaryCompare) = 0 Then
            txt_kappafree.Text = Trim(sel_dinterface_rs("TESTRESULT")) & ".00"
        Else
            txt_kappafree.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
        Else
            txt_kappafree.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
        
    ElseIf sel_dinterface_rs("TESTCODE") = "F042" Then
        'txt_lambdafree.Text = sel_dinterface_rs("TESTRESULT")
        If Trim(sel_dinterface_rs("TESTRESULT")) <> "" Then
        If InStr(1, Trim(sel_dinterface_rs("TESTRESULT")), ".", vbBinaryCompare) = 0 Then
            txt_lambdafree.Text = Trim(sel_dinterface_rs("TESTRESULT")) & ".00"
        Else
            txt_lambdafree.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
        Else
            txt_lambdafree.Text = Trim(sel_dinterface_rs("TESTRESULT"))
        End If
    End If
    sel_dinterface_rs.MoveNext
Wend
'***** Commented by Shraddha on 17/03/10 *****
'upd_dinterface_qr(0) = reqno_dinter
'upd_dinterface_qr.Execute
'*********************************************
End Sub

Private Sub cmb_reqno_Click()
'On Error GoTo msg:
VWREQNSTAT = ""
  Dim urvol As String
  urvol = ""
Dim PC As String
If Opt_patient.Value = True Then
    If NEWFLAG = True Then
            IMFLAG = False
            NEWFLAG = False
            ''****** ICU ****************
            VAR_REQNNO = req_chglen(cmb_reqno.Text)
            
            '''done by shrikant for display binterface data on screen
            Set rs_binterface = rdb.OpenResultset("SELECT * FROM interface.binterface WHERE REQNO='" & VAR_REQNNO & "' ", 1, 1)
            If rs_binterface.RowCount > 0 Then
                txt_iga = rs_binterface(3)
                txt_igg = rs_binterface(4)
                txt_igm = rs_binterface(5)
                txt_kappa = rs_binterface(6)
                txt_lambda = rs_binterface(7)
                
                
            End If
            '''''''''''''''''''''''''''''''''''''''''''''
            
            
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
            txt_age = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
            
                   Dim n As Integer
                   n = txt_age / 12
                   txt_age.Text = n
                       
               If Trim(txt_age.Text) = 0 Then
                  Dim age As String
                  age = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                  age = "." & (age Mod 12)
                     If age = ".0" Then
                        age = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                        age = ".0" & (age) 'Mod 30)
                     End If
                  txt_age.Text = Trim(age)
               End If
            
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
            
            If SELPATHDREQUISITNFALG1 = False Then
                SELPATHDREQUISITN_QR(0) = VAR_REQNNO ' req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            
           
             
            If selpathdrequisitn_rs.RowCount > 0 Then
                txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
                If IsNull(selpathdrequisitn_rs("OPPROCEDUR")) Then
                Else
                urvol = Trim(selpathdrequisitn_rs("OPPROCEDUR"))
                End If
            End If
            
             SERUM_FIELD
             SELDINTERFACE
             cmb_urrate.clear
             '***** SHWETA 07-NOV-12 *****
            cmb_urrate.AddItem (" ")
           ' cmb_urrate.AddItem ("Kappa  g/24hrs/" & urvol & " of urine.")
           ' cmb_urrate.AddItem ("Lambda g/24hrs/" & urvol & " of urine.")
            'cmb_urrate.AddItem ("Kappa is  Lambda is  gm/24hrs/" & urvol & "of urine.")
           ' cmb_urrate.AddItem ("Kappa  g/" & urvol & " Lambda  g/24hrs/" & urvol & " of urine.")
            '*****************************
            cmb_urrate.AddItem ("Kappa  g/24hrs urine")
            cmb_urrate.AddItem ("Lambda  g/24hrs urine")
            cmb_urrate.AddItem ("Kappa  g/24hrs & Lambda  g/24hrs urine")
             
            '***** Shraddha 08/12/09 *****
            dinter
            '*****************************
        
            If DMULPROCEFLAG = False Then
                DMULPROCE_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set DMULPROCE_RS = DMULPROCE_QR.OpenResultset(1, 3)
                DMULPROCEFLAG = True
            Else
                DMULPROCE_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                DMULPROCE_RS.MoreResults
                DMULPROCE_RS.Requery
            End If
        
        
        '''KANCHAN
        
'        If DMULPROCE_RS.RowCount > 0 Then
'
'                REPORT.Enabled = True
'                i = 1
'                MSFlexGrid1.Rows = 1
'                While Not DMULPROCE_RS.EOF
'
'
'                    If SELMA101FLAG = False Then
'                        SELMA101_QR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
'                        SELMA101FLAG = True
'                    Else
'                        SELMA101_QR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                        SELMA101_RS.MoreResults
'                        SELMA101_RS.Requery
'                    End If
'
'
'
'                    ''KANCHAN
'                        If SELDTEMPLETFLAG = False Then
'                            SELDTEMPLETQR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'                            SELDTEMPLETFLAG = True
'                        Else
'                            SELDTEMPLETQR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                            SELDTEMPLETRS.MoreResults
'                            SELDTEMPLETRS.Requery
'                        End If
'                        If SELDTEMPLETRS.RowCount > 0 Then
'                           RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
'                           PC = Trim(DMULPROCE_RS("pro_code"))
'                            If DMULPROCE_RS("pro_code") <> "F816" And DMULPROCE_RS("pro_code") <> "F828" And DMULPROCE_RS("pro_code") <> "F977" Then
'
'                                'INTERFACE 2005
'
'                                 If SELINTEREXPCDFLAG = False Then   'CHECKING FOR EXCETION VALUES
'                                     SELINTEREXPCDQR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                                     Set SELINTEREXPCDRS = SELINTEREXPCDQR.OpenResultset(1, 3)
'                                     SELINTEREXPCDFLAG = True
'                                 Else
'                                     SELINTEREXPCDQR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                                     SELINTEREXPCDRS.MoreResults
'                                     SELINTEREXPCDRS.Requery
'                                 End If
'                                 If SELINTEREXPCDRS.RowCount > 0 Then
'                                     MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & DMULPROCE_RS("pro_code") '& vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                 Else
'                                     'THIS IS FOR PICKUP TESTRESULT FROM INTERFACE/DINTERFACE
'
'                                      INTERFACEREAD_QR(1) = PC
'                                      INTERFACEREAD_QR(0) = VAR_REQNNO
'                                      Set INTERFACEREAD_RS = INTERFACEREAD_QR.OpenResultset(1, 1)
'
'                                      If INTERFACEREAD_RS.RowCount > 0 Then
'                                          INTFAC = INTERFACEREAD_RS("testresult")
'                                          If Trim(INTFAC) = "" Then
'                                              INTMODE = ""
'                                              MACHINENO = ""
'                                            Else
'                                              INTMODE = "I"
'                                              MACHINENO = INTERFACEREAD_RS("MACHINENO")
'                                          End If
'                                          MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(DMULPROCE_RS("PRO_CODE")) & vbTab & INTFAC & vbTab & RNAGE1 & vbTab & vbTab & INTMODE & vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL"))) & vbTab & MACHINENO
'                                      Else
'                                          INTFAC = ""
'                                          INTMODE = ""
'                                          MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & DMULPROCE_RS("pro_code") '& vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                      End If
'
'                                   End If
'                                'INTERFACE 2005
'
'                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                i = i + 1
'
'                            Else
'                                IMFLAG = True
'                                FRMIMM.Enabled = True
'                                cmb_elepattern.clear
'                                cmb_monband.clear
'                                cmb_locationnband.clear
'                                cmb_urineimm.clear
'                                CMB_BJPBJP.clear
'                                cmb_ka.clear
'                                TMADD
'
'                            End If
'
'                        End If
'                    'End If
'                    DMULPROCE_RS.MoveNext
'                Wend
                    
''''KANCHAN
'
'                    DMULPROCE_RS.MoveFirst
'                    While Not DMULPROCE_RS.EOF
'                    ''CHANGE
'                     If Trim(DMULPROCE_RS("PRO_CODE")) <> "F816" And Trim(DMULPROCE_RS("PRO_CODE")) <> "F828" And Trim(DMULPROCE_RS("PRO_CODE")) <> "F948" And Trim(DMULPROCE_RS("PRO_CODE")) <> "F977" Then
'                        'MSFlexGrid1.Row = 1
'                        'MSFlexGrid1.Col = 2
'                        txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'                        'txt_observed.Visible = True
'                        SSTab1.TabVisible(1) = False
'                    End If
'                        DMULPROCE_RS.MoveNext
'                    Wend
'                    If IMFLAG = True Then
'                        SSTab1.TabVisible(1) = True
'                    Else
'                        SSTab1.TabVisible(1) = False
'                    End If
'    End If

'''''''''''KANCHAN
'                If seldissecentrflag1 = False Then
'                    Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
'                    seldissecentrflag1 = True
'                Else
'                    seldissecentr1_rs.MoreResults
'                    seldissecentr1_rs.Requery
'                End If
'                If seldissecentr1_rs.RowCount > 0 Then
'                    cmb_enterby.clear
'                    While Not seldissecentr1_rs.EOF
'                        cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
'                        seldissecentr1_rs.MoveNext
'                    Wend
'                End If
'
'                If seldissecfinflag = False Then
'                    Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
'                    seldissecfinflag1 = True
'                Else
'                    seldissecfin1_rs.MoreResults
'                    seldissecfin1_rs.Requery
'                End If
'                If seldissecfin1_rs.RowCount > 0 Then
'                    cmb_finalby.clear
'                    While Not seldissecfin1_rs.EOF
'                        cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
'                        seldissecfin1_rs.MoveNext
'                    Wend
'                End If
''
    
    '******** code added by v.s.k. on 26-apr-2012 for drawing graph
    STRSQL = "select result from dislib.einterface where reqno='" & VAR_REQNNO & "'"
    Set RS = rdb.OpenResultset(STRSQL, 1, 1)
    If Not RS.EOF Then
        txtEInterface.Text = RS!Result
    End If
    RS.Close
    
    Picture1.AutoRedraw = True
    Picture1.Cls
    'Picture1.BackColor = vbRed
    K = 0
    For i = 1 To Len(txtEInterface.Text) Step 4
    
    ycor = Val("&H" & Mid(txtEInterface.Text, i, 4) & "&")
    
    Picture1.Line (K, Picture1.ScaleHeight)-(K, (Picture1.ScaleHeight - (ycor / 15))), vbBlack
    'Picture1.PSet (k, (Picture1.ScaleHeight - (ycor / 15))), vbBlack
    K = K + (Picture1.ScaleWidth / (Len(txtEInterface.Text) / 4))
    'MsgBox ycor
    Next
    Picture1.AutoRedraw = False

    
    '************* end
End If
      
      
      If viewflag = True Then
         viewflag = False
     '   txt_observed.Visible = False
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
                If REFDR_RS.EOF = False Then
                    cmb_enterby.Text = " " & Trim(REFDR_RS("enterby"))
                End If
                If REFDR_RS.EOF = False Then
                    cmb_finalby.Text = " " & Trim(REFDR_RS("freportby"))
                End If
                If REFDR_RS("COMMITFLAG") = "Y" Then
                    MsgBox "Report is commited , You can not Modify"
                    
                    ADDENDUMFALG = True
                    cmd_addendum.Visible = True
                    cmd_modify.Enabled = False
                     cmd_commit.Enabled = False
                    
                    
                    
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
      
      
        If selma104_rs.RowCount > 0 Then
            txt_ward.Visible = True
            txt_bedno.Visible = True
            lbl_ward.Visible = True
            lbl_bed.Visible = True
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
                SELPATHDREQUISITN_QR(0) = VAR_REQNNO ' req_chglen(Trim(cmb_reqno.Text))
                Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                SELPATHDREQUISITNFALG1 = True
            Else
                SELPATHDREQUISITN_QR(0) = VAR_REQNNO 'req_chglen(Trim(cmb_reqno.Text))
                selpathdrequisitn_rs.MoreResults
                selpathdrequisitn_rs.Requery
            End If
            If selpathdrequisitn_rs.RowCount > 0 Then
                txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
                
                If IsNull(selpathdrequisitn_rs("OPPROCEDUR")) Then
                Else
                urvol = Trim(selpathdrequisitn_rs("OPPROCEDUR"))
                End If
                
            End If
             cmb_urrate.AddItem (" ")
            'cmb_urrate.AddItem ("Kappa  g/24hrs/" & urvol & " of urine.")
            'cmb_urrate.AddItem ("Lambda g/24hrs/" & urvol & " of urine.")
            'cmb_urrate.AddItem ("Kappa is  Lambda is  gm/24hrs/" & urvol & "of urine.")
            'cmb_urrate.AddItem ("Kappa  g/" & urvol & " Lambda  g/24hrs/" & urvol & " of urine.")
            cmb_urrate.AddItem ("Kappa  g/24hrs urine")
            cmb_urrate.AddItem ("Lambda  g/24hrs urine")
            cmb_urrate.AddItem ("Kappa  g/24hrs & Lambda  g/24hrs urine")
            
           If DMPROCEDURFLAG = False Then
                 DMPROCEDUR_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set DMPROCEDUR_RS = DMPROCEDUR_QR.OpenResultset(1, 3)
                DMPROCEDURFLAG = True
            Else
                DMPROCEDUR_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                DMPROCEDUR_RS.MoreResults
                DMPROCEDUR_RS.Requery
            End If
            
            SERUM_FIELD
            '***** Shraddha 25/11/10 *****
            
            SELDIMBIOCHEQR(0) = req_chglen(cmb_reqno.Text)
            Set SELDIMBIOCHERS = SELDIMBIOCHEQR.OpenResultset(1, 3)
            
            If Not SELDIMBIOCHERS.EOF Then
                If InStr(1, Trim(SELDIMBIOCHERS("IGA")), ".", vbBinaryCompare) = 0 Then
                    txt_iga.Text = Trim(SELDIMBIOCHERS("IGA")) & ".00"
                Else
                    txt_iga.Text = Trim(SELDIMBIOCHERS("IGA"))
                End If
                
                If InStr(1, Trim(SELDIMBIOCHERS("IGM")), ".", vbBinaryCompare) = 0 Then
                    txt_igm.Text = Trim(SELDIMBIOCHERS("IGM")) & ".00"
                Else
                    txt_igm.Text = Trim(SELDIMBIOCHERS("IGM"))
                End If
                
                'txt_igm.Text = Trim(SELDIMBIOCHERS("IGM"))
                If InStr(1, Trim(SELDIMBIOCHERS("IGG")), ".", vbBinaryCompare) = 0 Then
                    txt_igg.Text = Trim(SELDIMBIOCHERS("IGG")) & ".00"
                Else
                    txt_igg.Text = Trim(SELDIMBIOCHERS("IGG"))
                End If
                
            End If
            '*****************************
            DMULPROCEFLAG = False
            If DMULPROCEFLAG = False Then
                DMULPROCE_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set DMULPROCE_RS = DMULPROCE_QR.OpenResultset(1, 3)
                DMULPROCEFLAG = True
            Else
                DMULPROCE_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                DMULPROCE_RS.MoreResults
                DMULPROCE_RS.Requery
            End If
            
            If DMULPROCE_RS.RowCount > 0 Then
              
                
                While Not DMULPROCE_RS.EOF
                If Trim(DMULPROCE_RS("pro_code")) = "F045" Or Trim(DMULPROCE_RS("pro_code")) = "F049" Or Trim(DMULPROCE_RS("pro_code")) = "F044" Or Trim(DMULPROCE_RS("pro_code")) = "F052" Or Trim(DMULPROCE_RS("pro_code")) = "F043" Or Trim(DMULPROCE_RS("pro_code")) = "F977" Or Trim(DMULPROCE_RS("pro_code")) = "F053" Or Trim(DMULPROCE_RS("pro_code")) = "F054" Then  'Trim(DMULPROCE_RS("pro_code")) = "F816"
                              
                    REPORT.Enabled = False
                    FRMIMM.Enabled = False
                    If SELDIMBIOCHEFLAG = False Then
                        SELDIMBIOCHEQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                        Set SELDIMBIOCHERS = SELDIMBIOCHEQR.OpenResultset(1, 3)
                        SELDIMBIOCHEFLAG = True
                    Else
                        SELDIMBIOCHEQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                        SELDIMBIOCHERS.MoreResults
                        SELDIMBIOCHERS.Requery
                    End If
                    
                    If SELDIMBIOCHERS.RowCount > 0 Then '29/06/10
                        IMFLAG = True
                        SSTab1.TabVisible(1) = True
                        cmb_elepattern.Text = Trim(SELDIMBIOCHERS("ELEPHPATT"))
                        cmb_monband.Text = Trim(SELDIMBIOCHERS("MONBAND"))
                        cmb_locationnband = Trim(SELDIMBIOCHERS("LOCNBAND"))
                        If SELDIMBIOCHERS.EOF = False Then
                        txt_connband = " " & Trim(SELDIMBIOCHERS("CONMBAND"))
                        End If
'                        txt_iga.Text = Trim(SELDIMBIOCHERS("IGA"))
'                        txt_igm.Text = Trim(SELDIMBIOCHERS("IGM"))
'                        txt_igg.Text = Trim(SELDIMBIOCHERS("IGG"))
                        If SELDIMBIOCHERS.EOF = False Then
                            txt_remmbond.Text = " " & SELDIMBIOCHERS("REMMBAND").Value
                        Else
                            txt_remmbond.Text = " "
                        End If
                        cmb_urineimm.Text = SELDIMBIOCHERS("UIMELPHO")
                        CMB_BJPBJP.Text = SELDIMBIOCHERS("BJP")
                        '************** Shraddha Pandit 23/01/09 ****************
                        SEL_BIOCHE1_QR(0) = req_chglen(cmb_reqno.Text)
                        Set SEL_BIOCHE1_RS = SEL_BIOCHE1_QR.OpenResultset(1, 3)
                        If Not SEL_BIOCHE1_RS.EOF Then
                            If SEL_BIOCHE1_RS("BJPCOMMENT") <> "" Or SEL_BIOCHE1_RS("BJPCOMMENT") <> Null Then
                                txt_urrate.Text = Trim(SEL_BIOCHE1_RS("BJPCOMMENT"))
                                '************** Shweta 07-11-2012 ****************
                                cmb_urrate.Text = Trim(SEL_BIOCHE1_RS("BJPCOMMENT"))
                                '**************
                            End If
                            
                           '********* Added by Shraddha Pandit on 24/04/09 ********
                            If SEL_BIOCHE1_RS("KAPPAUR") <> "" Or SEL_BIOCHE1_RS("KAPPAUR") <> Null Then
                                txt_urkappa.Text = Trim(SEL_BIOCHE1_RS("KAPPAUR"))
                            End If
                            
                            If SEL_BIOCHE1_RS("LAMBDAUR") <> "" Or SEL_BIOCHE1_RS("LAMBDAUR") <> Null Then
                                txt_urlambda.Text = Trim(SEL_BIOCHE1_RS("LAMBDAUR"))
                            End If
                            '******************************************************
                        End If
                        '**********************************************************
                        
                        cmb_ka.Text = SELDIMBIOCHERS("KALA")
                        If InStr(1, Trim(SELDIMBIOCHERS("KAPPA")), ".", vbBinaryCompare) = 0 Then
                            txt_kappa.Text = Trim(SELDIMBIOCHERS("KAPPA")) & ".00"
                        Else
                            txt_kappa.Text = Trim(SELDIMBIOCHERS("KAPPA"))
                        End If
                        
                        If InStr(1, Trim(SELDIMBIOCHERS("LAMBDA")), ".", vbBinaryCompare) = 0 Then
                            txt_kappa.Text = Trim(SELDIMBIOCHERS("LAMBDA")) & ".00"
                        Else
                            txt_lambda.Text = " " & Trim(SELDIMBIOCHERS("LAMBDA"))
                        End If
                        
                        
                        '***** Added by Shraddha on 08/12/09 *****
                        SEL_BIOCHE1_QR(0) = req_chglen(cmb_reqno.Text)
                        Set SEL_BIOCHE1_RS = SEL_BIOCHE1_QR.OpenResultset(1, 3)
                        If Not SEL_BIOCHE1_RS.EOF Then
                            If Trim(SEL_BIOCHE1_RS("KAPPAFREE")) <> "" Then '11/12/09
                                If InStr(1, Trim(SEL_BIOCHE1_RS("KAPPAFREE")), ".", vbBinaryCompare) = 0 Then
                                    txt_kappafree.Text = Trim(SEL_BIOCHE1_RS("KAPPAFREE")) & ".00"
                                Else
                                    txt_kappafree.Text = Trim(SEL_BIOCHE1_RS("KAPPAFREE"))
                                End If
                            End If
                            If Trim(SEL_BIOCHE1_RS("LAMBDAFREE")) <> "" Then '11/12/09
                                
                                If InStr(1, Trim(SEL_BIOCHE1_RS("LAMBDAFREE")), ".", vbBinaryCompare) = 0 Then
                            
                                    txt_lambdafree.Text = Trim(SEL_BIOCHE1_RS("LAMBDAFREE")) & ".00"
                                Else
                                    txt_lambdafree.Text = Trim(SEL_BIOCHE1_RS("LAMBDAFREE"))
                                End If
                            End If
                        End If
                        '*****************************************
                        If SELDIMBIOCHERS.EOF = False Then
                            txt_monoclonal.Text = " " & SELDIMBIOCHERS("TMONPIMFIX").Value
                        Else
                            txt_monoclonal.Text = " "
                        End If
                    End If
                'Else
                '    IMFLAG = False
                End If
                DMULPROCE_RS.MoveNext
                
              Wend
            End If
            
            If IMFLAG = True Then
                 SSTab1.TabVisible(1) = True
            Else
                 SSTab1.TabVisible(1) = False
            End If
        If DMPROCEDUR_RS.RowCount > 0 Then
                REPORT.Enabled = False
                FRMIMM.Enabled = False
                i = 1
                MSFlexGrid1.Rows = 1
                DMPROCEDUR_RS.MoveFirst
                While Not DMPROCEDUR_RS.EOF
                If SELMA101FLAG = False Then
                        SELMA101_QR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                        SELMA101FLAG = True
                    Else
                        SELMA101_QR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                        SELMA101_RS.MoreResults
                        SELMA101_RS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                    
                        '***** Shraddha 07/04/10 *****
                        'SELECT * FROM interface.dinterface where reqno=? and TESTCODE=?
                        SEL_DINT_QR(0) = Trim(req_chglen(cmb_reqno.Text))
                        SEL_DINT_QR(1) = Trim(DMPROCEDUR_RS("PRO_CODE"))
                        Set SEL_DINT_RS = SEL_DINT_QR.OpenResultset(1, 3)
                        'SELDTEMPLETQR(1) = Trim(SEL_DINT_RS("MACHINENO"))
                        '*****************************
                        
                        str = "SELECT * FROM DISLIB.DTEMPLET1 WHERE PRO_CODE=? AND MACHINNA=?"
                        Set SELDTEMPLETQR = rdb.CreateQuery("", str)
                        On Error Resume Next
                        'If SELDTEMPLETFLAG = False Then
                            SELDTEMPLETQR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
                            'SELDTEMPLETQR(1) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)) '07/04/10
                            
                            SELDTEMPLETQR(1) = Trim(SEL_DINT_RS("MACHINENO"))
                            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
                         '   SELDTEMPLETFLAG = True
                        'Else
                         '   SELDTEMPLETQR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
                            'SELDTEMPLETQR(1) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)) '07/04/10
                          '  SELDTEMPLETQR(1) = Trim(SEL_DINT_RS("MACHINENO"))
                           ' SELDTEMPLETRS.MoreResults
                            'SELDTEMPLETRS.Requery
                       ' End If
                        If SELDTEMPLETRS.RowCount > 0 Then
                            tmflag = True
                            'MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(DMPROCEDUR_RS("PRO_CODE")) & vbTab & DMPROCEDUR_RS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & Trim(DMPROCEDUR_RS("pvalue")) & vbTab & Trim(DMPROCEDUR_RS("source")) & vbTab & vbTab & Trim(DMPROCEDUR_RS("entryby"))
                            
                            'MSFlexGrid1.AddItem I & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(DMPROCEDUR_RS("PRO_CODE")) & vbTab & DMPROCEDUR_RS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & Trim(DMPROCEDUR_RS("pvalue")) & vbTab & Trim(DMPROCEDUR_RS("source")) & vbTab & vbTab & Trim(DMPROCEDUR_RS("REMARKS"))
                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(DMPROCEDUR_RS("PRO_CODE")) & vbTab & DMPROCEDUR_RS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & Trim(DMPROCEDUR_RS("pvalue")) & vbTab & Trim(DMPROCEDUR_RS("source")) & vbTab & vbTab & Trim(DMPROCEDUR_RS("REMARKS"))
                            '***** Shraddha 05/11/10 *****
                             If Trim(DMPROCEDUR_RS("PRO_CODE")) = "F821" Then
                                '**** Age *****
                                If txt_age.Text >= 1 And txt_age.Text <= 60 Then
                                    RNAGE = 60
                                End If
                                If (txt_age.Text >= 61 And txt_age.Text <= 100) Then
                                    RNAGE = 100
                                End If
                                '*************
                                SELDBCALKPHSQR(0) = "F821"
                                SELDBCALKPHSQR(1) = RNAGE 'Trim(txt_age.Text)
                                SELDBCALKPHSQR(2) = Trim(DMPROCEDUR_RS("REMARKS")) 'SELINTER_RS("MACHINENO") 'Trim(DMPROCEDUR_RS("REMARKS"))
                                Set SELDBCALKPHSRS = SELDBCALKPHSQR.OpenResultset(1, 3)
                                If Not SELDBCALKPHSRS.EOF Then
                                   
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = Trim(SELDBCALKPHSRS!range)
                                End If
                            End If
                            '*****************************
                            If Trim(DMPROCEDUR_RS("PRO_CODE")) = "F817" Then
                                If Val(DMPROCEDUR_RS("OBSERVALUE")) > 200 Then
                                  chk_var = True
                                Else
                                    chk_var = False
                                End If
                            End If
                            
                            i = i + 1
                        End If
                        
                    End If
                DMPROCEDUR_RS.MoveNext
                Wend
           
            End If
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
                txt_comments.Text = " " & SELDMBCCOMMRS("COMMENTS")
            End If
      End If

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
      
End If

'***** Shraddha 05/12/2011 *****
'For showing previous values
Call prev_val
'*******************************

If opt_staff.Value = True Then
    STAFF
End If

'***** Added by Shraddha *****

If Trim(txt_kappa.Text) <> "" And Trim(txt_lambda.Text) <> "" Then
Dim rto As Double
 txt_kappa.Text = Replace(Trim(txt_kappa.Text), "^", "")
 txt_lambda.Text = Replace(Trim(txt_lambda.Text), "^", "") 'Added on 13/05/10 by Shraddha
 txt_lambda.Text = Replace(Trim(txt_lambda.Text), "<", "")
 
 'txt_kappa.Text = Replace(Trim(txt_kappa.Text), ".", "")
 
    rto = CInt(Trim(txt_kappa.Text)) / CInt(Trim(txt_lambda.Text))
    
    MsgBox "Kappa/Lambda is : " & rto
    'Uncommented and added again on 07/04/10
    
    'Commented and added by Shraddha on 30/03/10
'    txt_lambda.Text = Replace(Trim(txt_lambda.Text), "^", "")
'    rto = Trim(txt_lambda.Text) / Trim(txt_kappa.Text)
'    MsgBox "Lambda/Kappa is : " & rto
End If
'*****************************************
''***** Shraddha 02/01/2011 *****
'    Dim SEL_STPRINT_QR As rdoQuery
'    Dim SEL_STPRINT_RS As rdoResultset
'
'    str = "SELECT * FROM DISLIB.STOP_PRINT WHERE MODALITY='TM'"
'    Set SEL_STPRINT_QR = rdb.CreateQuery("", str)
'
'    'SEL_STPRINT_QR(0) = Trim(txt_modcode.Text)
'    Set SEL_STPRINT_RS = SEL_STPRINT_QR.OpenResultset(1, 3)
'
'    'REQNO = req_chglen(Trim(cmb_reqno.Text))
'
'    If Not SEL_STPRINT_RS.EOF Then
'        If SEL_STPRINT_RS("FLAG") = "Y" Then
'            cmd_print.Enabled = False
'        Else
'            cmd_print.Enabled = True
'        End If
'    End If
'    '*******************************
'***** Shraddha 22/09/2012 *****
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
    If REFDR_RS("COMMITFLAG") = "Y" Then
        'MsgBox "Report is commited , You can not Modify"
    cmd_modify.Enabled = False
    cmd_commit.Enabled = False
    End If
    '*******************************
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

Public Sub SERUM_FIELD()

            If DMULPROCEFLAG = False Then
                DMULPROCE_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                Set DMULPROCE_RS = DMULPROCE_QR.OpenResultset(1, 3)
                DMULPROCEFLAG = True
            Else
                DMULPROCE_QR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
                DMULPROCE_RS.MoreResults
                DMULPROCE_RS.Requery
            End If
            
            FRMIMM.Enabled = True
If DMULPROCE_RS.RowCount > 0 Then
        
            IMFLAG = True
            
            Frame2.Visible = False
            Frame6.Visible = False
            Frame4.Visible = False
            Frame5.Visible = False
                txt_iga.Enabled = False
                txt_igm.Enabled = False
                txt_igg.Enabled = False
                txt_kappa.Enabled = False
                txt_lambda.Enabled = False
                txt_kappafree.Enabled = False
                txt_lambdafree.Enabled = False
                lbl_typmono.Visible = False
                txt_monoclonal.Visible = False
                txt_monoclonal.Enabled = True
                CMB_BJPBJP.Enabled = False
                txt_urrate.Enabled = False
                txt_urkappa.Enabled = False
                txt_urlambda.Enabled = False
            
        While Not DMULPROCE_RS.EOF
            
            If Trim(DMULPROCE_RS("pro_code")) = "F043" Then
            
                Frame2.Visible = True
                Frame6.Visible = True
                Frame4.Visible = True
                Frame5.Visible = False
                txt_iga.Enabled = True
                txt_igm.Enabled = True
                txt_igg.Enabled = True
                txt_kappa.Enabled = True
                txt_lambda.Enabled = True
                txt_kappafree.Enabled = True
                txt_lambdafree.Enabled = True
                lbl_typmono.Visible = True
                txt_monoclonal.Visible = True
                txt_monoclonal.Enabled = True
            End If
            
            
            ''serum Immunoglobins
            
            If Trim(DMULPROCE_RS("pro_code")) = "F045" Then
                Frame2.Visible = True
                Frame2.Enabled = True
                txt_iga.Enabled = True
                txt_igm.Enabled = True
                txt_igg.Enabled = True
                
            ElseIf Trim(DMULPROCE_RS("pro_code")) = "F046" Then
                Frame2.Visible = True
                Frame2.Enabled = True
                txt_iga.Enabled = True
            ElseIf Trim(DMULPROCE_RS("pro_code")) = "F047" Then
                Frame2.Visible = True
                Frame2.Enabled = True
                txt_igm.Enabled = True
            ElseIf Trim(DMULPROCE_RS("pro_code")) = "F048" Then
                Frame2.Visible = True
                Frame2.Enabled = True
                txt_igg.Enabled = True
            End If
            
            ''''IFE
            If Trim(DMULPROCE_RS("pro_code")) = "F052" Then
                lbl_typmono.Visible = True
                txt_monoclonal.Visible = True
                txt_monoclonal.Enabled = True
               
            End If
            
            ''' Serum Light Chains
            
            If Trim(DMULPROCE_RS("pro_code")) = "F049" Then
                Frame6.Visible = True
                Frame6.Enabled = True
                txt_kappa.Enabled = True
                txt_lambda.Enabled = True
                txt_kappafree.Enabled = True
                txt_lambdafree.Enabled = True
            ElseIf Trim(DMULPROCE_RS("pro_code")) = "F050" Then
                Frame6.Visible = True
                Frame6.Enabled = True
                txt_kappa.Enabled = True
            ElseIf Trim(DMULPROCE_RS("pro_code")) = "F051" Then
                Frame6.Visible = True
                Frame6.Enabled = True
                txt_lambda.Enabled = True
            ElseIf Trim(DMULPROCE_RS("pro_code")) = "F041" Then
                Frame6.Visible = True
                Frame6.Enabled = True
                txt_kappafree.Enabled = True
            ElseIf Trim(DMULPROCE_RS("pro_code")) = "F042" Then
                Frame6.Visible = True
                Frame6.Enabled = True
                txt_lambdafree.Enabled = True
            End If
            
            ''''SPE
            If Trim(DMULPROCE_RS("pro_code")) = "F044" Then
                Frame4.Enabled = True
                Frame4.Visible = True
            End If
            
            ''''BJP
            If Trim(DMULPROCE_RS("pro_code")) = "F053" And Trim(DMULPROCE_RS("pro_code")) = "F054" And Trim(DMULPROCE_RS("pro_code")) = "F977" Then
                    Frame2.Visible = False
                    Frame6.Visible = False
                    Frame4.Visible = False
                    Frame5.Visible = True
                    Frame5.Enabled = True
                    CMB_BJPBJP.Enabled = False
                    txt_urrate.Enabled = False
                    txt_urkappa.Enabled = False
                    txt_urlambda.Enabled = False
                    FRMIMM.Caption = ""
             Else
                    FRMIMM.Caption = "Complete Serum Protein Electrophoresis (SPE) Profile"
            End If
            
            If Trim(DMULPROCE_RS("pro_code")) = "F977" Then
                FRMIMM.Enabled = True
                Frame2.Visible = False
                Frame6.Visible = False
                Frame4.Visible = False
                Frame5.Visible = True
                Frame5.Enabled = True
                
                CMB_BJPBJP.Enabled = True
                txt_urrate.Enabled = True
                FRMIMM.Caption = ""

            ElseIf Trim(DMULPROCE_RS("pro_code")) = "F053" Then
                    FRMIMM.Enabled = True
                    Frame2.Visible = False
                    Frame6.Visible = False
                    Frame4.Visible = False
                    Frame5.Visible = True
                    Frame5.Enabled = True
                 txt_urkappa.Enabled = True
                 FRMIMM.Caption = ""

            ElseIf Trim(DMULPROCE_RS("pro_code")) = "F054" Then
                    FRMIMM.Enabled = True
                    Frame2.Visible = False
                    Frame6.Visible = False
                    Frame4.Visible = False
                    Frame5.Visible = True
                    Frame5.Enabled = True
                txt_urlambda.Enabled = True
                FRMIMM.Caption = ""

            End If
            
            DMULPROCE_RS.MoveNext
        Wend

End If


End Sub

Private Sub cmb_urrate_Change()
txt_urrate.Text = cmb_urrate.Text
End Sub

Private Sub cmb_urrate_Click()
txt_urrate.Text = ""
'SendKeys "{HOME}"
End Sub

Private Sub cmd_addendum_Click()
  cmd_save.Enabled = True
  txt_addendum.Text = ""
End Sub

Private Sub cmd_back_Click()
btnenabl ("INI")
clear
txt_observed.Visible = False
REPORT.Enabled = False
FRMIMM.Enabled = False
IMFLAG = False
tmflag = False
txt_iga.BackColor = vbWhite
txt_igm.BackColor = vbWhite
txt_igg.BackColor = vbWhite
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_cancel_Click()
btnenabl ("INI")
clear
cmb_machine.Visible = False
txt_observed.Visible = False
REPORT.Enabled = False
FRMIMM.Enabled = False
IMFLAG = False
tmflag = False
txt_iga.BackColor = vbWhite
txt_igm.BackColor = vbWhite
txt_igg.BackColor = vbWhite
End Sub

Private Sub cmd_commit_Click()
Dim i As Integer
'If Len(Trim(cmb_finalby)) = 0 Then
'    m = MsgBox("The Finalized by field is EMPTY.", vbCritical)
'    cmb_finalby.SetFocus
'    Exit Sub
'End If
'******** Added by Shraddha Pandit on 11/02/2009 **********
If chk_var = True Then
    MsgBox "Please check Dilution level", vbInformation, " Tumour Marker "
End If
'**********************************************************
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
    updrrequisitn1_qr(0) = UCase(Left(cmb_finalby.Text, 30))
    updrrequisitn1_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = req_chglen(cmb_reqno.Text)
    updrrequisitn1_qr.Execute
    
    MsgBox "Record is Committed Successfully."
    
    'RTIMELOG 18-01-05
                 UPDRTIMEC_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                 UPDRTIMEC_QR(1) = txt_time.Text
                 UPDRTIMEC_QR(2) = req_chglen(cmb_reqno.Text)
                 UPDRTIMEC_QR.Execute
    
End If
rdb.CommitTrans
btnenabl ("INI")

cmd_print.Enabled = True
''***** Shraddha 02/01/2011 *****
'    Dim SEL_STPRINT_QR As rdoQuery
'    Dim SEL_STPRINT_RS As rdoResultset
'
'    str = "SELECT * FROM DISLIB.STOP_PRINT WHERE MODALITY='TM'"
'    Set SEL_STPRINT_QR = rdb.CreateQuery("", str)
'
'    'SEL_STPRINT_QR(0) = Trim(txt_modcode.Text)
'    Set SEL_STPRINT_RS = SEL_STPRINT_QR.OpenResultset(1, 3)
'
'    'REQNO = req_chglen(Trim(cmb_reqno.Text))
'
'    If Not SEL_STPRINT_RS.EOF Then
'        If SEL_STPRINT_RS("FLAG") = "Y" Then
'            cmd_print.Enabled = False
'        Else
'            cmd_print.Enabled = True
'        End If
'    End If
'    '*******************************
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
'***************
If Len(Trim(MaskEdBox1.Text)) > 0 Then
    VAR1 = MaskEdBox1.FormattedText
    Dim strLink As String
    'strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
    strLink = "http://10.80.1.9/emrtmhact/TOPBAR.ASP?CASE_NO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If

'***** 20/04/11 *****
'
'Public Sub OpenWebsite(URL As String)
'Dim ret&
'ret& = ShellExecute(Me.hwnd, "Open", URL, vbNullString, vbNullString, SW_SHOW)
End Sub

Private Sub cmd_modify_Click()
Dim str As String
btnenabl ("NEW")
FLDENABL ("MODIFY")
'txt_observed.Visible = True
txt_observed.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
MODIFYFLAG = True
 TMADD

flagsave = 2
REPORT.Enabled = True
FRMIMM.Enabled = True
MSFlexGrid1.col = 2

str = cmb_elepattern.Text
cmb_elepattern.clear
cmb_elepattern.AddItem "Normal"
cmb_elepattern.AddItem "Abnormal"
cmb_elepattern.Text = str
str = ""
str = cmb_monband.Text
cmb_monband.clear
cmb_monband.AddItem "Present"
cmb_monband.AddItem "Not Seen"
cmb_monband.Text = str
str = ""
str = cmb_locationnband.Text
cmb_locationnband.clear
cmb_locationnband.AddItem "Gamma"
cmb_locationnband.AddItem "Beta"
cmb_locationnband.AddItem "Alpha"
cmb_locationnband.AddItem "Gamma+Beta"
cmb_locationnband.AddItem "Gamma+Alpha"
cmb_locationnband.AddItem "Alpha+beta"
cmb_locationnband.Text = str
str = ""
str = cmb_urineimm.Text
cmb_urineimm.clear
cmb_urineimm.AddItem "Normal"
cmb_urineimm.AddItem "Ab.Normal"
cmb_urineimm.Text = str
str = ""
str = CMB_BJPBJP.Text
CMB_BJPBJP.clear
    CMB_BJPBJP.AddItem "Positive"
    CMB_BJPBJP.AddItem "Negative"
    CMB_BJPBJP.AddItem "Not Requested"
CMB_BJPBJP.Text = str
str = ""
str = cmb_ka.Text
cmb_ka.clear
cmb_ka.AddItem "Kappa"
cmb_ka.AddItem "Lamda"
cmb_ka.Text = str

'
'            If seldissecentrflag1 = False Then
'                    Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
'                    seldissecentrflag1 = True
'                Else
'                    seldissecentr1_rs.MoreResults
'                    seldissecentr1_rs.Requery
'                End If
'                If seldissecentr1_rs.RowCount > 0 Then
'                    STR = ""
'                    STR = cmb_enterby.Text
'                    cmb_enterby.clear
'                    While Not seldissecentr1_rs.EOF
'                        cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
'                        seldissecentr1_rs.MoveNext
'                    Wend
'                    cmb_enterby.Text = STR
'                End If
'
'                If seldissecfinflag = False Then
'                    Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
'                    seldissecfinflag1 = True
'                Else
'                    seldissecfin1_rs.MoreResults
'                    seldissecfin1_rs.Requery
'                End If
'                If seldissecfin1_rs.RowCount > 0 Then
'                    STR = ""
'                    STR = cmb_finalby.Text
'                    cmb_finalby.clear
'                    While Not seldissecfin1_rs.EOF
'                        cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
'                        seldissecfin1_rs.MoveNext
'                    Wend
'                    cmb_finalby.Text = STR
'                End If

'changes made by amarjeet on 08-11-2002 for  commit logic to disable txt_observed
If ADDENDUMFALG Then 'means commited
txt_observed.Visible = False
Else
txt_observed.Width = MSFlexGrid1.CellWidth
txt_observed.Height = MSFlexGrid1.CellHeight
txt_observed.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
txt_observed.Visible = True
End If
txt_comments.Enabled = True
'commented by amarjeet prev logic
'txt_observed.Width = MSFlexGrid1.CellWidth
'txt_observed.Height = MSFlexGrid1.CellHeight
'txt_observed.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
'txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
End Sub
Private Sub cmd_new_Click()
VWREQNSTAT = ""
Dim L1 As Integer
btnenabl ("NEW")
FLDENABL ("ENA")
'txt_observed.Visible = True
clear
IMFLAG = False
tmflag = False
txt_iga.BackColor = vbWhite
txt_igm.BackColor = vbWhite
txt_igg.BackColor = vbWhite
    If DATETIMEFLAG = False Then
    datetime
    DATETIMEFLAG = True
     Else
    datetime
    End If
    txt_date.Text = DATE11
    txt_time.Text = time1
    cmb_enterby.Text = cmb_enterby.List(0)
    cmb_enterby.Locked = True
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
cmb_dept.SetFocus
Opt_patient.Value = True
MaskEdBox1.Enabled = True
NEWFLAG = True
txt_observed.Visible = False
flagsave = 1
End Sub

Private Sub cmd_print_Click()

Dim n, d As String

'On Error GoTo MSG:

'RTIMELOG 18-01-05
    UPDRTIMEP_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    UPDRTIMEP_QR(1) = txt_time.Text
    UPDRTIMEP_QR(2) = req_chglen(cmb_reqno.Text)
    UPDRTIMEP_QR.Execute

Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim rto
Dim i As Integer
 'On Error GoTo prtmsg:

DoEvents
 
 
'''***** Added by Shraddha on 17/03/10 *****
'If Trim(txt_kappa.Text) <> "" And Trim(txt_lambda.Text) <> "" Then
'    rto = Trim(txt_kappa.Text) / Trim(txt_lambda.Text)
'    MsgBox "Kappa/Lambda is : " & rto
'End If
'''*****************************************


'******************************* for displaying designation ************'
'********** added by milind on 03-03-06
If Trim(cmb_finalby.Text) = "" Then
    MsgBox "Please enter Final by"
    Exit Sub
End If

userid_qr(0) = Trim(cmb_finalby.Text)
Set userid_rs = userid_qr.OpenResultset(1, 3)
n = userid_rs(0)
userdesc_qr(0) = n
Set userdesc_rs = userdesc_qr.OpenResultset(1, 3)
d = userdesc_rs(0)
'************upto here**********

Dim VAR11 As Integer

For VAR11 = 0 To 12
    CrystalReport1.Formulas(VAR11) = ""
Next VAR11

VAR_REQNNO = req_chglen(cmb_reqno.Text)

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
            If DMULPROCEFLAG = False Then
                DMULPROCE_QR(0) = req_chglen(cmb_reqno.Text)
                Set DMULPROCE_RS = DMULPROCE_QR.OpenResultset(1, 3)
                DMULPROCEFLAG = True
            Else
                DMULPROCE_QR(0) = req_chglen(cmb_reqno.Text)
                DMULPROCE_RS.MoreResults
                DMULPROCE_RS.Requery
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
'        If seldrequisitn_rs.RowCount > 0 Then
'            prepath = ""
'
'            While Not seldrequisitn_rs.EOF
'                If Trim(seldrequisitn_rs("pathno")) <> Trim(txt_pathno.Text) Then
'                    prepath = prepath & seldrequisitn_rs("pathno")
'                End If
'                seldrequisitn_rs.MoveNext
'            Wend
'
'        End If

 str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
  'COMMENTED AND ADDED BY SHWETA 24-SEPT-2010
' CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
 CrystalReport1.SelectionFormula = "{DISBCUNILF.REQNO} = '" & VAR_REQNNO & "'"
 If Opt_patient.Value = True Then
    If tmflag = True Then
     'COMMENTED AND ADDED BY SHWETA 24-SEPT-2010
            'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM.RPT"
            CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISTMFORBATCH.rpt"
            '***** Added by Shraddha on 26/03/10 *****
'            'CR_SubTM.SelectionFormula = "({RREQUISITN.CASENO} = '" & MaskEdBox1.FormattedText & "') AND ({RREQUISITN.REQNNO} LIKE 'FECTM%')"
'            CR_SubTM.SelectionFormula = "{RREQUISITN.CASENO} = '" & MaskEdBox1.FormattedText & "'"
'            CR_SubTM.ReportFileName = "C:\DISREPORTSEC\SubTM.RPT"
'
'
'            CR_SubTM.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'
'            'CR_SubTM.SelectionFormula = "{MOLECULAB.REQNO} = '" & req_chglen(cmb_reqno.Text) & "'"
'
'            'UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)
'
'
'            CR_SubTM.WindowState = crptMaximized
'            CR_SubTM.Action = 1
            '*****************************************
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
        'Commented by shraddha on 08/08/2011
'        If Not SELMEMO_RS.EOF Then
'            If IsNull(SELMEMO_RS("MEMONO")) And Trim(SELMEMO_RS("PAT_CATG")) <> "NC" Then
'
'                '***** Shraddha 08/08/2011 *****
'                For S = 0 To 25
'                    CrystalReport1.Formulas(S) = ""
'                Next S
'                '*******************************
'
'                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTMICU.RPT"
'                'SHWETA 01-OCT-2010
'                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'            End If
'        End If
' till here 17/03/10
    End If
            
            
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
    End If
        
        'justin
        
'        If SELDIMBIOCHEPFLAG = False Then
'            SELDIMBIOCHEPQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
'            Set SELDIMBIOCHEPRS = SELDIMBIOCHEPQR.OpenResultset(1, 3)
'            SELDIMBIOCHEPFLAG = True
'        Else
'            SELDIMBIOCHEPQR(0) = VAR_REQNNO 'req_chglen(cmb_reqno.Text)
'            SELDIMBIOCHEPRS.MoreResults
'            SELDIMBIOCHEPRS.Requery
'        End If
'
'        If SELDIMBIOCHEPRS("COMMITFLAG") = "Y" Then
        
        ''''' Done By Kanchan
        'If Trim(cmb_elepattern.Text) <> "" And SSTab1.TabVisible(1) = True Then
                If SSTab1.TabVisible(1) = True Then
            
        'justin
                
        'If IMFLAG = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM1.RPT"
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        If Trim(prepath) = "" Then
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
        End If
        
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "SERUM  ELECTROPHORESIS  REPORT" & "'"
            str1 = ""
            str1 = "Entered by " & Trim(cmb_enterby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "SERUM  ELECTROPHORESIS  REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(17) = "DESG = '" & d & "'"
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    End If
End If
If opt_staff.Value = True Then
        If tmflag = True Then
     
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTMS.RPT"
            CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
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
            '''ADDED BY KANCHAN
            CLPTM
            CrystalReport1.Formulas(20) = "COLLAPS TIME='" & CLPTIME & "'"
            
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
            CrystalReport1.Action = 1
        End If
        
        
        'If IMFLAG = True Then
        'Added next or instead of and on 15/12/08 by Shraddha
        If Trim(cmb_elepattern.Text) <> "" Or SSTab1.TabVisible(1) = True Then
        

        
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM1S.RPT"
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
        'CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISTMFORBATCH.rpt"
        
        '25/11/10
        
        VAR11 = 0
        For VAR11 = 0 To 50
            CrystalReport1.Formulas(VAR11) = ""
        Next VAR11
        
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
        If Trim(prepath) = "" Then
            CrystalReport1.Formulas(1) = "PREPATH = ''"
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
        
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(10) = "FRSEC = ''"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    End If
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


Dim macerror As Boolean
    If SSTab1.TabVisible(0) = True And SSTab1.TabVisible(1) = True Then
        MsgBox " Pls check whether Serum Electrophoresis Data is Entered or not !!!!", vbInformation, " Tumor Marker "
    End If
'If SSTab1.Tab = 1 Then
'    If txt_urrate.Text = "" Then
'    MsgBox "Urine Excreate rate is not entered", vbCritical
'    cmb_urrate.SetFocus
'    Exit Sub
'    End If
'End If
    macerror = False
    For i = 1 To MSFlexGrid1.Rows - 1
      If Len(Trim(MSFlexGrid1.TextMatrix(i, 1))) > 0 Then
      
        If Len(Trim(MSFlexGrid1.TextMatrix(i, 7))) = 0 Then
         macerror = True
        End If
      End If
      
    Next i
    
    If macerror = True Then
       MsgBox "Machine Name not filled in the last column"
       Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    assign
''



'''
    
'
'           UPDINTERFACE_QR(0) = 3
'           UPDINTERFACE_QR(1) = " " 'Trim(MSFlexGrid1.TextMatrix(i, 7))
'           UPDINTERFACE_QR(2) = req_chglen(cmb_reqno.Text)
'           UPDINTERFACE_QR(3) = sc
'           UPDINTERFACE_QR.Execute
'
    btnenabl ("SAVE")
    Me.MousePointer = vbNormal
    cmd_print.Enabled = True
'    '***** Shraddha 02/01/2011 *****
'    Dim SEL_STPRINT_QR As rdoQuery
'    Dim SEL_STPRINT_RS As rdoResultset
'
'    str = "SELECT * FROM DISLIB.STOP_PRINT WHERE MODALITY='TM'"
'    Set SEL_STPRINT_QR = rdb.CreateQuery("", str)
'
'    'SEL_STPRINT_QR(0) = Trim(txt_modcode.Text)
'    Set SEL_STPRINT_RS = SEL_STPRINT_QR.OpenResultset(1, 3)
'
'    'REQNO = req_chglen(Trim(cmb_reqno.Text))
'
'    If Not SEL_STPRINT_RS.EOF Then
'        If SEL_STPRINT_RS("FLAG") = "Y" Then
'            cmd_print.Enabled = False
'        Else
'            cmd_print.Enabled = True
'        End If
'    End If
'    '*******************************
    'cmd_commit.Enabled = True
    
    cmd_cancel.Enabled = True
End Sub

Private Sub cmd_view_Click()
btnenabl ("VIEW")
FLDENABL ("VIEW")
txt_observed.Visible = False
clear
    If DATETIMEFLAG = False Then
    datetime
    DATETIMEFLAG = True
     Else
    datetime
    End If
    
    txt_date.Text = DATE11
    txt_time.Text = time1
    IMFLAG = False
    tmflag = False
    txt_iga.BackColor = vbWhite
    txt_igm.BackColor = vbWhite
    txt_igg.BackColor = vbWhite
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
End Sub








Private Sub MSFlexGrid1_Click()

If MSFlexGrid1.col = 2 Then
        txt_observed = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))
        txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        txt_observed.Visible = True
End If

If MSFlexGrid1.col = 7 Then
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
   ' MaskEdBox1.Mask = "#######"  ' COMMENTED BY VARINDER
End If
End Sub
Private Sub Form_Load()
On Error GoTo MSG:
'******* Shraddha 11/02/09 *******
chk_var = False
'*********************************
cmb_machine.Visible = False
txt_observed.Visible = False
cmd_addendum.Visible = False
FLDENABL ("DIS")

'********************displaying designation in crystal report*********'
'******************added by milind on 03-03-06
str = "select usr_id from dislib.dislnkusr where username=?"
Set userid_qr = rdb.CreateQuery("userid_qr", str)

str = "select user_des from dislib.dsecurity where usr_id=?"
Set userdesc_qr = rdb.CreateQuery("userdesc_qr", str)

'***********SHWETA 07-NOV-2011  ****************'
str = "select TESTNAME from dislib.FECCOMMENT WHERE MODALITY='TM'"
Set COMMENTLIST_qr = rdb.CreateQuery("COMMLST", str)
Set COMMENTLIST_rs = COMMENTLIST_qr.OpenResultset(1, 3)
txt_comments.clear
If COMMENTLIST_rs.EOF = False Then
COMMENTLIST_rs.MoveFirst
While Not COMMENTLIST_rs.EOF
txt_comments.AddItem (Trim(COMMENTLIST_rs(0)))
COMMENTLIST_rs.MoveNext
Wend

End If
'*******************************'
'-----------

'*******************************'
'************upto here*********************

If DATETIMEFLAG = False Then
    datetime
    DATETIMEFLAG = True
     Else
    datetime
    End If
    txt_date.Text = DATE11
    txt_time.Text = time1
If TMREPORTFLAG = False Then
    TMREPORTFLAG = True
   
   

   
'*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1TM", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    cmb_enterby.AddItem LOGINUSRNAME
    cmb_enterby.Text = (LOGINUSRNAME)
    cmb_enterby.Text = cmb_enterby.List(0)
    cmb_enterby.Locked = True
Else
    MsgBox "The Security Details of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE='FEC' AND USRTYPE>=8 ORDER BY USERNAME "
Set REPWTRN_QR = rdb.CreateQuery("REPWTRN_STRAA", str)
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
        End If
        REPWTRNRS.MoveNext
    Wend
    cmb_finalby.Text = cmb_finalby.List(0)
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11TM", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)

'STR = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE=?"
'Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1TM", STR)
'SELFINALBY_QR(0) = 6
'Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)
'If Not SELFINALBY_RS.EOF Then
'    Do While Not SELFINALBY_RS.EOF
'       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
'       SELFINALBY_RS.MoveNext
'    Loop
'End If


'*********SECURITY**********************************

''********ICU***************
str = "SELECT ICDS2 FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELICU_QR = rdb.CreateQuery("SELICU_QR11111", str)

str = "SELECT MEMONO,PAT_CATG FROM RISLIB.RREQUISITN WHERE REQNNO=?" ' AND (PAT_CATG<>'NC' OR PAT_CATG<>'SF')"
Set SELMEMO_QR = rdb.CreateQuery("SELMEMO_QR11111", str)


''********ICU***************

'***** Shraddha 05/01/11 *****
str = "SELECT * FROM DISLIB.DBCALKPHS1 WHERE PROCODE=? AND AGE2=? AND MACHINNA=? "
Set SELDBCALKPHSQR = rdb.CreateQuery("", str)
'*****************************

str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ')"
Set sel_deptqr = rdb.CreateQuery("sel1_deptqrTUMREP1111", str)

'for retriving reqnos
Set req_qr = rdb.CreateQuery("1reqtumrep", "select * from rislib.RREQUISITN where caseno=? AND REQNSTAT='E' AND MODALITYCD=? ORDER BY REQNDT desc ")



' for modality selection
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel1_moadlityqr11tumrep", str)

' for retriveing name,age etc
Set SELTA101_QR = rdb.CreateQuery("selta101_qr11TUMREP", "select * from PATOOLSLIB.ta101 where case_no =?")

  ' for bedno, ward retriveing
   str = "select * from PATOOLSLIB.ma104 where case_no =?"
   Set selma104_qr = rdb.CreateQuery("selma1041_qr11TUMREP", str)
   
   
' FOR REFEFING DOCTOR
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("REFDR1_QR1TUMREP", str)
' FOR REFEFING DOCTOR
    str = "SELECT pro_code FROM DISLIB.DMULPROCE WHERE REQNO=? "
    Set DMULPROCE_QR = rdb.CreateQuery("DMULPROCE_QR1TUMREP", str)
    
    str = "SELECT * FROM DISLIB.DMPROCEDUR WHERE REQNO=? "
    Set DMPROCEDUR_QR = rdb.CreateQuery("DMPROCEDUR_QR1TUMREP", str)
    
'FOR INSERTING
    'STR = "insert into  dislib.DMPROCEDUR(reqno,PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,PVALUE) values (?,?,?,?,?,?,?)"
    'STR = "insert into  dislib.DMPROCEDUR(reqno,PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,PVALUE,source,entryby) values (?,?,?,?,?,?,?,?,?)"
    ' THE MACHINE NO IS SHIFTED TO REMARKS COLUMN INSTEAD OF ENTRYBY AS REMARKS IS NOT USED TILL NOW BY ANYBODY
    str = "insert into  dislib.DMPROCEDUR(reqno,PRO_CODE,OBSERVALUE,OBSER_TYPE,UOM,NRANGE,PVALUE,source,REMARKS) values (?,?,?,?,?,?,?,?,?)"
    Set INSDPRO = rdb.CreateQuery("INSDPRO1TUMREP", str)
    
    
'FOR UPDATING
  str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE REQNNO=?"
  Set updrrequisitn_qr = rdb.CreateQuery("updrrequisitn_qr1BC", str)
   
  str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD=? AND DEPT_CODE=? AND (REQNSTAT=? OR REQNSTAT=?)ORDER BY REQNDT DESC "
  Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR11TUMREP", str)
    
 'STR = "UPDATE DISLIB.DMPROCEDUR SET OBSERVALUE=?,OBSER_TYPE=?,UOM=?,NRANGE=?,PVALUE=? WHERE REQNO = ? AND PRO_CODE =?"
  str = "UPDATE DISLIB.DMPROCEDUR SET OBSERVALUE=?,OBSER_TYPE=?,UOM=?,NRANGE=?,PVALUE=?,SOURCE=?,ENTRYBY=? WHERE REQNO = ? AND PRO_CODE =?"
  Set UPDDMPROCEDUR_QR = rdb.CreateQuery("UPDDMPROCEDUR_QR1TUMREPTM", str)

'FOR RETRIEVING THE PROCEDURE NAME
    'STR = "SELECT * FROM DISLIB.MA101 WHERE SER_CODE = ?"
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE = ?"
  Set SELMA101_QR = rdb.CreateQuery(SELMA101_QRTUMREP, str)
    
'FOR RETRIEVING THE PROCEDURE CODE
  'STR = "SELECT * FROM DISLIB.MA101 WHERE SER_DET = ?"
  str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_DET = ?"
  Set SELMA101_QR1 = rdb.CreateQuery(SELMA101_QR1TUMREP, str)
  
  str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? AND EMPCODE=?"
  Set SELDREQSTAFFD1QR1 = rdb.CreateQuery("SELDREQSTAFFD1QR11TUMREP", str)
   
  '***** Shraddha 07/04/10 *****
  str = "SELECT * FROM interface.dinterface where reqno=? and TESTCODE=?"
  Set SEL_DINT_QR = rdb.CreateQuery("", str)
  
  '***** 28/07/10 *****
  str = "UPDATE interface.dinterface SET MACHINENO=? WHERE REQNO=? and TESTCODE=?"
  Set UPD_DInterface = rdb.CreateQuery("", str)
  
  '*****************************
    
    
  'str = "SELECT * FROM DISLIB.DTEMPLET WHERE PRO_CODE=? "
  'Commented and added by Shraddha on 07/04/10
  str = "SELECT * FROM DISLIB.DTEMPLET1 WHERE PRO_CODE=? AND MACHINNA=?"
  Set SELDTEMPLETQR = rdb.CreateQuery("SELDTEMPLETQRTM", str)
  
  
  
  
  
  str = "SELECT * FROM DISLIB.DTEMPLET1 WHERE PRO_CODE=? AND MACHINNA=?"
  Set SELDTEMPLETQR1 = rdb.CreateQuery("SELDTEMPLETQRTM1", str)
   
  
  str = "SELECT * FROM DISLIB.DIMBIOCHEL WHERE REQNO=?"
  Set SELDIMBIOCHEQR = rdb.CreateQuery("SELDIMBIOCHEQRTM", str)
  
  'justin
  str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
  Set SELDIMBIOCHEPQR = rdb.CreateQuery("SELDIMBIOCHEQRTMp", str)
  
  str = "INSERT INTO DISLIB.DIMBIOCHEL(REQNO,ELEPHPATT,MONBAND,LOCNBAND,CONMBAND,IGA,IGM,IGG,REMMBAND,UIMELPHO,BJP,KALA,TMONPIMFIX,KAPPA,LAMBDA) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
  'STR = "INSERT INTO DISLIB.DIMBIOCHEL(REQNO,ELEPHPATT,MONBAND,LOCNBAND,CONMBAND,IGA,IGM,IGG,REMMBAND,UIMELPHO,BJP,KALA,TMONPIMFIX,polyclonal) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
  Set INSDIMBIOCHEQR = rdb.CreateQuery("INSDIMBIOCHEQRTM", str)
  
  str = "INSERT INTO DISLIB.DIMBIOCHE1(REQNO,KAPPAFREE,LAMBDAFREE,BJPCOMMENT,KAPPAUR,LAMBDAUR) VALUES (?,?,?,?,?,?)"
  Set INSDIMBIOCHE1 = rdb.CreateQuery("INSERTDIMBIOCHE1", str)
  
  str = "UPDATE DISLIB.DIMBIOCHEL SET ELEPHPATT=?,MONBAND=?,LOCNBAND=?,CONMBAND=?,IGA=?,IGM=?,IGG=?,REMMBAND=?,UIMELPHO=?,BJP=?,KALA=?,TMONPIMFIX=?,KAPPA=?,LAMBDA=? WHERE REQNO=?"
  'STR = "UPDATE DISLIB.DIMBIOCHEL SET ELEPHPATT=?,MONBAND=?,LOCNBAND=?,CONMBAND=?,IGA=?,IGM=?,IGG=?,REMMBAND=?,UIMELPHO=?,BJP=?,KALA=?,TMONPIMFIX=?,POLYCLONAL=? WHERE REQNO=?"
  Set UPDDIMBIOCHEQR = rdb.CreateQuery("UPDDIMBIOCHEQRTM", str)
  
  str = "UPDATE DISLIB.DIMBIOCHE1 SET KAPPAFREE=?,LAMBDAFREE=?,BJPCOMMENT=?,KAPPAUR=?,LAMBDAUR=? WHERE REQNO=?"
  Set UPDDIMBIOCHE1 = rdb.CreateQuery("UPDATEDIMBIOCHE1", str)
  
  
  '*************** Shraddha Pandit 23/01/09 ***************
  str = "SELECT * FROM DISLIB.DIMBIOCHE1 WHERE REQNO=?"
  Set SEL_BIOCHE1_QR = rdb.CreateQuery("", str)
  '********************************************************
  
  
  str = "INSERT INTO DISLIB.DMBCCOMM(REQNO,COMMENTS) VALUES(?,?)"
  Set INSDMBCCOMMQR = rdb.CreateQuery("INSDMBCCOMMQRBtm", str)
  
  str = "UPDATE DISLIB.DMBCCOMM SET COMMENTS=? WHERE REQNO=?"
  Set UPDDMBCCOMMQR = rdb.CreateQuery("UPDDMBCCOMMQRtm", str)
  
  str = "SELECT * FROM DISLIB.DMBCCOMM WHERE REQNO=?"
  Set SELDMBCCOMMQR = rdb.CreateQuery("SELDMBCCOMMQRtm", str)
    
   str = "select * from dislib.dissec where enterby='YES'AND MODALITY='TM'"
   Set seldissecentr1_qr = rdb.CreateQuery("seldissecentr1_qr1TM", str)
   
   str = "select * from dislib.dissec where FINALby='YES'AND MODALITY='TM'"
   Set seldissecfin1_qr = rdb.CreateQuery("seldissecFIN1_qr1TM", str)
   
   str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
   Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn111_qrTM", str)
  
  str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
  Set SELPATHDREQUISITN_QR = rdb.CreateQuery("SELPATHDREQUISITN1_QRTM", str)
  
 '''KANCHAN
    str = "SELECT * FROM INTERFACE.DINTERFACE WHERE REQNO=?"
    Set SELINTER_QR = rdb.CreateQuery("CGSDF", str)
    
    '***** Shraddha 08/12/09 *****
    str = "SELECT * FROM INTERFACE.DINTERFACE WHERE REQNO=? and (iflag = '1' OR iflag = '2')"
    Set sel_dinterface_qr = rdb.CreateQuery("", str)
    
    str = "update INTERFACE.dinterface set iflag = '3' where reqno = ?"
    Set upd_dinterface_qr = rdb.CreateQuery("", str)
    '*****************************
    
    str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
    Set selMA101QR = rdb.CreateQuery("SELMA101QRBC", str)
 
 'THIS IS RTIMLOG
 
 str = "update rislib.rtimelog set ADPREPDATE=?,ADPREPTIME=? where reqnno=?"
 Set UPDRTIMECADD_QR = rdb.CreateQuery("updatetimecadd1tm", str)
 
 str = "UPDATE RISLIB.RTIMELOG SET SAVREPDATE=?,SAVREPTIME=? WHERE REQNNO=?"
 Set UPDRTIME_QR = rdb.CreateQuery("QWEQWEQWEtm", str)

 str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEC_QR = rdb.CreateQuery("QWEQWEQWEctm", str)
 
 
 str = "UPDATE RISLIB.RTIMELOG SET PRNREPDATE=?,PRNREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEP_QR = rdb.CreateQuery("QWEQWEQWEPtm", str)

 str = "UPDATE RISLIB.RTIMELOG SET MODREPDATE=?,MODREPTIME=? WHERE REQNNO=?"
 Set UPDRTIMEM_QR = rdb.CreateQuery("QWEQWEQWEMtm", str)

'THIS IS RTIMLOG
  
'ADDENDUM
  str = "UPDATE DISLIB.ADDENDUMCY SET ADDENDNO=?,ADDEND_DT=?,ADDEND_TM=?,USER_ID=?,ADD_MAX_LN=? WHERE REQNO=?"
  Set updateaddendum_qr = rdb.CreateQuery("UPDATEADDENDUMtm", str)
  
  str = "INSERT INTO DISLIB.ADDENDUMCY VALUES(?,?,?,?,?,?)"
  Set insertaddendum_qr = rdb.CreateQuery("INSERTADDENDUMtm", str)
  
  str = "SELECT * FROM DISLIB.ADDENDUMCY WHERE REQNO=?"
  Set selectaddendum_qr = rdb.CreateQuery("SELECTADDENDtm", str)
  
'INTERFACE
 
  str = "SELECT EXCPCODE FROM DISLIB.INTEREXPCD WHERE EXCPCODE=?"
  Set SELINTEREXPCDQR = rdb.CreateQuery("SELECTEXPCODETM", str)
  SELINTEREXPCDFLAG = False

  str = "SELECT * FROM INTERFACE.DINTERFACE WHERE IFLAG = '2' AND REQNO = ? AND TESTCODE = ?"
  Set INTERFACEREAD_QR = rdb.CreateQuery("INTERFACEDATATM1", str)

  str = "UPDATE INTERFACE.DINTERFACE SET IFLAG=?,MACHINENO=? WHERE REQNO=? AND TESTCODE=?"
  Set UPDINTERFACE_QR = rdb.CreateQuery("UPDINTER", str)

End If

REPORT.Enabled = False
FRMIMM.Enabled = False
MSFlexGrid1.col = 2
txt_observed.Visible = False
txt_observed.Width = MSFlexGrid1.CellWidth
txt_observed.Height = MSFlexGrid1.CellHeight
txt_observed.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
btnenabl ("INI")
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
        If Right(Trim(cmb_dept.List(ind_cbodept)), 3) = "FEC" Then
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
cmb_elepattern.Text = ""
cmb_locationnband.Text = ""
cmb_monband.Text = ""
txt_connband.Text = ""
txt_remmbond.Text = ""
cmb_urineimm.Text = ""
CMB_BJPBJP.Text = ""
cmb_ka.Text = ""
txt_iga.Text = ""
txt_igm.Text = ""
txt_igg.Text = ""
txt_kappa.Text = ""
txt_lambda.Text = ""
txt_comments.Text = ""
txt_monoclonal.Text = ""
txt_kappafree.Text = ""
txt_lambdafree.Text = ""
txt_urkappa.Text = ""
txt_urlambda.Text = ""
txt_urrate.Text = ""
cmb_polyclonal.clear
cmb_urrate.clear
'cmb_urrate.Text = ""
'For i = 1 To MSFlexGrid1.Row
'For J = 0 To MSFlexGrid1.Col
'MSFlexGrid1.TextMatrix(i, J) = " "
'Next J
'Next i
MSFlexGrid1.Rows = 1
txt_observed.Text = " "
txtEInterface.Text = ""
Picture1.Cls
NEWFLAG = False
viewflag = False
End Sub

Private Sub MaskEdBox1_LostFocus()
Dim CASENO As String
'On Error GoTo msg:
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
                req_qr(0) = Trim(MaskEdBox1.FormattedText)
                req_qr(1) = Trim(txt_modcode.Text)
                Set req_rs = req_qr.OpenResultset(1, 3)
                REQFLAG = True
          Else
                req_qr(0) = Trim(MaskEdBox1.FormattedText)
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
                
                SELRREQUISITN_QR(0) = Trim(MaskEdBox1.FormattedText)
                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELRREQUISITN_QR(3) = "T"
                SELRREQUISITN_QR(4) = "F"
                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                SELRREQUISITNFLAG = True
            Else
                SELRREQUISITN_QR(0) = Trim(MaskEdBox1.FormattedText)
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


End Sub

Private Sub Timer1_Timer()

If Val(txt_iga.Text) > 450 Then
    If i <> 1000 Then
       TXT_IG.Visible = True
       txt_iga.Visible = False
      
       i = Timer1.Interval
       j = j + 1
       
    Else
        TXT_IG.Visible = False
        txt_iga.Visible = True
        i = 0
        j = j + 1
       If j > 10 Then
            Timer1.Interval = 0
            TXT_IG.Visible = False
            txt_iga.Enabled = True
            Exit Sub
       End If
    End If
End If

End Sub

Private Sub Timer2_Timer()

If Val(txt_igm.Text) > 250 Then
    If i <> 1000 Then
       TXT_IG.Visible = True
       txt_igm.Visible = False

       i = Timer2.Interval
       K = K + 1

    Else
        TXT_IG.Visible = False
        txt_igm.Visible = True
        i = 0
        K = K + 1
       If K > 10 Then
            Timer2.Interval = 0
            Exit Sub
       End If
    End If
End If
End Sub

Private Sub Timer3_Timer()
If Val(txt_igg.Text) > 1500 Then
            If i <> 1000 Then
               TXT_IGG1.Visible = True
               txt_igg.Visible = False
              
               i = Timer3.Interval
               l = l + 1
               
            Else
                TXT_IGG1.Visible = False
                txt_igg.Visible = True
                i = 0
                l = l + 1
                If l > 10 Then
                     Timer3.Interval = 0
                     Exit Sub
                End If
            End If
        End If
End Sub



Private Sub txt_iga_LostFocus()
'If Trim(txt_iga.Text) >= 90 Or Trim(txt_iga.Text) <= 450 Then
'txt_iga.BackColor = vbWhite
'End If
'
'If Trim(txt_iga.Text) > 450 Then
'
'    txt_iga.BackColor = RGB(255, 100, 120)
''    For i = 0 To 1000
''        Beep
''    Next i
'End If
'If Trim(txt_iga.Text) < 90 Then
' txt_iga.BackColor = RGB(100, 255, 150)
'' For i = 0 To 1000
''    Beep
'' Next i
'End If
If InStr(1, Trim(txt_iga.Text), ".", vbBinaryCompare) = 0 Then
    txt_iga.Text = Trim(txt_iga.Text) & ".00"
Else
    txt_iga.Text = txt_iga.Text
End If
End Sub

Private Sub txt_igg_LostFocus()
'If Trim(txt_igg.Text) = "" Then
'Exit Sub
'End If
'    If Trim(txt_igg.Text) >= 700 Or Trim(txt_igg.Text) <= 1500 Then
'    txt_igg.BackColor = vbWhite
'    End If
'
'If Trim(txt_igg.Text) > 1500 Then
'
'    txt_igg.BackColor = RGB(255, 100, 120)
''    For i = 0 To 1000
''        Beep
''    Next i
'
'End If
'If Trim(txt_igg.Text) < 700 Then
' txt_igg.BackColor = RGB(100, 255, 150)
'' For i = 0 To 1000
''    Beep
'' Next i
'End If
If InStr(1, Trim(txt_igg.Text), ".", vbBinaryCompare) = 0 Then
    txt_igg.Text = Trim(txt_igg.Text) & ".00"
Else
    txt_igg.Text = Trim(txt_igg.Text)
End If
End Sub

Private Sub txt_igm_LostFocus()
'If Trim(txt_igm.Text) = "" Then
'    Exit Sub
'End If
'If Trim(txt_igm.Text) >= 60 Or Trim(txt_igm.Text) <= 250 Then
'txt_igm.BackColor = vbWhite
'End If
'Timer2.Interval = 1000
'If Trim(txt_igm.Text) > 250 Then
'
'    txt_igm.BackColor = RGB(255, 100, 120)
''    For i = 0 To 1000
''        Beep
''    Next i
'
'End If
'If Trim(txt_igm.Text) < 60 Then
' txt_igm.BackColor = RGB(100, 255, 150)
'' For i = 0 To 1000
''    Beep
'' Next i
'End If
If InStr(1, Trim(txt_igm.Text), ".", vbBinaryCompare) = 0 Then
    txt_igm.Text = Trim(txt_igm.Text) & ".00"
Else
    txt_igm.Text = Trim(txt_igm.Text)
End If
End Sub



Private Sub txt_kappa_LostFocus()
If InStr(1, Trim(txt_kappa.Text), ".", vbBinaryCompare) = 0 Then
    txt_kappa.Text = Trim(txt_kappa.Text) & ".00"
Else
    txt_kappa.Text = Trim(txt_kappa.Text)
End If
End Sub

Private Sub txt_kappafree_LostFocus()
If InStr(1, Trim(txt_kappafree.Text), ".", vbBinaryCompare) = 0 Then
    txt_kappafree.Text = Trim(txt_kappafree.Text) & ".00"
Else
    txt_kappafree.Text = Trim(txt_kappafree.Text)
End If
End Sub

Private Sub txt_lambda_LostFocus()
If InStr(1, Trim(txt_lambda.Text), ".", vbBinaryCompare) = 0 Then
    txt_lambda.Text = Trim(txt_lambda.Text) & ".00"
Else
    txt_lambda.Text = Trim(txt_lambda.Text)
End If
End Sub



Private Sub txt_lambdafree_LostFocus()
If InStr(1, Trim(txt_lambdafree.Text), ".", vbBinaryCompare) = 0 Then
    txt_lambdafree.Text = Trim(txt_lambdafree.Text) & ".00"
Else
    txt_lambdafree.Text = Trim(txt_lambdafree.Text)
End If
End Sub

Private Sub txt_observed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    '*** color coding
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
'        If SELDTEMPLETFLAG = False Then
'            SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
'            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'            SELDTEMPLETFLAG = True
'        Else
'            SELDTEMPLETQR(0) = Trim(SELDMA101RS("SER_CODE"))
'            SELDTEMPLETRS.MoreResults
'            SELDTEMPLETRS.Requery
'        End If
'
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
'                    MSFlexGrid1.CellBackColor = RGB(255, 255, 255)
'
'                 End If
'            End If
'        End If
'    End If


    '*** color coding
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
'    i = MSFlexGrid1.Rows - 1
'    If i = MSFlexGrid1.Row Then
'        cmd_save.SetFocus
'        Exit Sub
'    End If
'     MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = txt_observed.Text
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
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2) = txt_observed.Text
End Sub


Public Sub STAFF()
 SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
    Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
    If selpathdrequisitn_rs.RowCount > 0 Then
        txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
    End If
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
    '*****SHWETA 11-05-2012 AS TOLD BY KALPANA MAM TO DISABLE MODIFY BUTTON IF REPORT IS COMMITED
  VWREQNSTAT = REFDR_RS("REQNSTAT")
  If VWREQNSTAT = "F" Then
  cmd_modify.Enabled = False
  cmd_commit.Enabled = False
  Else
  cmd_modify.Enabled = True
  cmd_commit.Enabled = True
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
    
    '****
    
    '''' KANCHAN
'          If DMULPROCEFLAG = False Then
'                DMULPROCE_QR(0) = req_chglen(cmb_reqno.Text)
'                Set DMULPROCE_RS = DMULPROCE_QR.OpenResultset(1, 3)
'                DMULPROCEFLAG = True
'            Else
'                DMULPROCE_QR(0) = req_chglen(cmb_reqno.Text)
'                DMULPROCE_RS.MoreResults
'                DMULPROCE_RS.Requery
'            End If
'
'
'         If DMULPROCE_RS.RowCount > 0 Then
'
'                REPORT.Enabled = True
'                i = 1
'                MSFlexGrid1.Rows = 1
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
'                    If SELMA101_RS.RowCount > 0 Then
'                        If SELDTEMPLETFLAG = False Then
'                            SELDTEMPLETQR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                            'SELDTEMPLETQR1(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
'                            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
'                            SELDTEMPLETFLAG = True
'                        Else
'                            SELDTEMPLETQR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                            SELDTEMPLETRS.MoreResults
'                            SELDTEMPLETRS.Requery
'                        End If
'                        RNAGE1 = Trim(SELDTEMPLETRS!LOWVALUE) & "-" & Trim(SELDTEMPLETRS!HIGHVALUE)
'                        If SELDTEMPLETRS.RowCount > 0 Then
'                           PC = Trim(DMULPROCE_RS("pro_code"))
'                            If DMULPROCE_RS("pro_code") <> "F816" And DMULPROCE_RS("pro_code") <> "F828" And DMULPROCE_RS("pro_code") <> "F977" Then
'
'                               'INTERFACE 2005
'                                If SELINTEREXPCDFLAG = False Then   'CHECKING FOR EXCETION VALUES
'                                     SELINTEREXPCDQR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                                     Set SELINTEREXPCDRS = SELINTEREXPCDQR.OpenResultset(1, 3)
'                                     SELINTEREXPCDFLAG = True
'                                 Else
'                                     SELINTEREXPCDQR(0) = Trim(DMULPROCE_RS("PRO_CODE"))
'                                     SELINTEREXPCDRS.MoreResults
'                                     SELINTEREXPCDRS.Requery
'                                 End If
'                                 If SELINTEREXPCDRS.RowCount > 0 Then
'                                     MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & DMULPROCE_RS("pro_code") '& vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                 Else
'                                     'THIS IS FOR PICKUP TESTRESULT FROM INTERFACE/DINTERFACE
'
'                                      INTERFACEREAD_QR(1) = PC
'                                      INTERFACEREAD_QR(0) = VAR_REQNNO
'                                      Set INTERFACEREAD_RS = INTERFACEREAD_QR.OpenResultset(1, 1)
'
'                                      If INTERFACEREAD_RS.RowCount > 0 Then
'                                          INTFAC = INTERFACEREAD_RS("testresult")
'                                          If Trim(INTFAC) = "" Then
'                                              INTMODE = ""
'                                              MACHINENO = ""
'                                            Else
'                                              INTMODE = "I"
'                                              MACHINENO = INTERFACEREAD_RS("MACHINENO")
'                                          End If
'                                          MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(DMULPROCE_RS("PRO_CODE")) & vbTab & INTFAC & vbTab & RNAGE1 & vbTab & vbTab & INTMODE & vbTab & getpreval(Trim(INTERFACEREAD_RS("PREVAL"))) & vbTab & MACHINENO
'                                      Else
'                                          INTFAC = ""
'                                          INTMODE = ""
'                                          MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & DMULPROCE_RS("pro_code") '& vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                      End If
'
'                                 End If
'                                'INTERFACE 2005
'
''                                MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & vbTab & " " & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE"))
'                                i = i + 1
'                            Else
'                                IMFLAG = True
'                                FRMIMM.Enabled = True
'                                cmb_elepattern.clear
'                                cmb_monband.clear
'                                cmb_locationnband.clear
'                                cmb_urineimm.clear
'                                CMB_BJPBJP.clear
'                                cmb_ka.clear
'                                TMADD
''                                If DMULPROCE_RS("pro_code") = "F828" Or DMULPROCE_RS("pro_code") = "F816" Then
''                                    Frame4.Enabled = True
''                                End If
'                            End If
'                        End If
'                    End If
'                    DMULPROCE_RS.MoveNext
'                Wend
                    
               SERUM_FIELD 'Shraddha 25/11/10
               SELDINTERFACE
                    
'                    DMULPROCE_RS.MoveFirst
'                    While Not DMULPROCE_RS.EOF
'                     If Trim(DMULPROCE_RS("PRO_CODE")) <> "F816" And Trim(DMULPROCE_RS("PRO_CODE")) <> "F828" And Trim(DMULPROCE_RS("PRO_CODE")) <> "F977" Then
'                        'MSFlexGrid1.Row = 1
'                        'MSFlexGrid1.Col = 2
'                        txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'                        'txt_observed.Visible = True
'                        SSTab1.TabVisible(1) = False 'here
'                     End If
'                        DMULPROCE_RS.MoveNext
'                    Wend
                    If IMFLAG = True Then
                        SSTab1.TabVisible(1) = True
                    Else
                        SSTab1.TabVisible(1) = False 'here1
                       ' FRMIMM.Enabled = True
                    End If
                
                     txt_observed.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                    'txt_observed.Visible = True

        'End If ''KANCHAN
'                If seldissecentrflag1 = False Then
'                    Set seldissecentr1_rs = seldissecentr1_qr.OpenResultset(1, 3)
'                    seldissecentrflag1 = True
'                Else
'                    seldissecentr1_rs.MoreResults
'                    seldissecentr1_rs.Requery
'                End If
'                If seldissecentr1_rs.RowCount > 0 Then
'                    cmb_enterby.clear
'                    While Not seldissecentr1_rs.EOF
'                        cmb_enterby.AddItem Trim(seldissecentr1_rs("CCNAME"))
'                        seldissecentr1_rs.MoveNext
'                    Wend
'                End If
'
'                If seldissecfinflag = False Then
'                    Set seldissecfin1_rs = seldissecfin1_qr.OpenResultset(1, 3)
'                    seldissecfinflag1 = True
'                Else
'                    seldissecfin1_rs.MoreResults
'                    seldissecfin1_rs.Requery
'                End If
'                If seldissecfin1_rs.RowCount > 0 Then
'                    cmb_finalby.clear
'                    While Not seldissecfin1_rs.EOF
'                        cmb_finalby.AddItem Trim(seldissecfin1_rs("CCNAME"))
'                        seldissecfin1_rs.MoveNext
'                    Wend
'                End If
'
    
    '****
    
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
                If Trim(SELDREQSTAFFD1RS1("RELATION")) <> "SELF" Then
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
        '****
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
                DMULPROCE_QR(0) = req_chglen(cmb_reqno.Text)
                Set DMULPROCE_RS = DMULPROCE_QR.OpenResultset(1, 3)
                DMULPROCEFLAG = True
            Else
                DMULPROCE_QR(0) = req_chglen(cmb_reqno.Text)
                DMULPROCE_RS.MoreResults
                DMULPROCE_RS.Requery
            End If
            
           SERUM_FIELD 'Shraddha 25/11/10
            
            If DMULPROCE_RS.RowCount > 0 Then
                While Not DMULPROCE_RS.EOF
                If Trim(DMULPROCE_RS("pro_code")) = "F816" Then
'
'                    REPORT.Enabled = False
'                    FRMIMM.Enabled = False
                    If SELDIMBIOCHEFLAG = False Then
                        SELDIMBIOCHEQR(0) = req_chglen(cmb_reqno.Text)
                        Set SELDIMBIOCHERS = SELDIMBIOCHEQR.OpenResultset(1, 3)
                        SELDIMBIOCHEFLAG = True
                    Else
                        SELDIMBIOCHEQR(0) = req_chglen(cmb_reqno.Text)
                        SELDIMBIOCHERS.MoreResults
                        SELDIMBIOCHERS.Requery
                    End If
                    If SELDIMBIOCHERS.RowCount > 0 Then
                        IMFLAG = True
                        cmb_elepattern.Text = Trim(SELDIMBIOCHERS("ELEPHPATT"))
                        cmb_monband.Text = Trim(SELDIMBIOCHERS("MONBAND"))
                        cmb_locationnband = Trim(SELDIMBIOCHERS("LOCNBAND"))
                            If SELDIMBIOCHERS.EOF = False Then
                                txt_connband = " " & Trim(SELDIMBIOCHERS("CONMBAND"))
                            End If
                        txt_iga.Text = Trim(SELDIMBIOCHERS("IGA"))
                        txt_igm.Text = Trim(SELDIMBIOCHERS("IGM"))
                        txt_igg.Text = Trim(SELDIMBIOCHERS("IGG"))
                            If SELDIMBIOCHERS.EOF = False Then
                                txt_remmbond.Text = " " & SELDIMBIOCHERS("REMMBAND").Value
                            Else
                                txt_remmbond.Text = " "
                            End If
                        cmb_urineimm.Text = SELDIMBIOCHERS("UIMELPHO")
                        CMB_BJPBJP.Text = SELDIMBIOCHERS("BJP")
                        cmb_ka.Text = SELDIMBIOCHERS("KALA")
                        txt_kappa.Text = SELDIMBIOCHERS("KAPPA")
                        txt_lambda.Text = SELDIMBIOCHERS("LAMBDA")
                        
                            If SELDIMBIOCHERS.EOF = False Then
                                txt_monoclonal.Text = " " & SELDIMBIOCHERS("TMONPIMFIX").Value
                            Else
                                txt_monoclonal.Text = " "
                            End If
                    End If
                    
                End If
                DMULPROCE_RS.MoveNext
              Wend
            End If
            
            '***** Shraddha 25/11/10 *****
            
            SELDIMBIOCHEQR(0) = req_chglen(cmb_reqno.Text)
            Set SELDIMBIOCHERS = SELDIMBIOCHEQR.OpenResultset(1, 3)
        
            If Not SELDIMBIOCHERS.EOF Then
                txt_iga.Text = Trim(SELDIMBIOCHERS("IGA"))
                txt_igm.Text = Trim(SELDIMBIOCHERS("IGM"))
                txt_igg.Text = Trim(SELDIMBIOCHERS("IGG"))
                CMB_BJPBJP.Text = Trim(SELDIMBIOCHERS("BJP"))
            End If
            
            '*****************************
            '***** Shraddha 27/06/2011 alpha *****
            SEL_BIOCHE1_QR(0) = req_chglen(cmb_reqno.Text)
            Set SEL_BIOCHE1_RS = SEL_BIOCHE1_QR.OpenResultset(1, 3)
            If Not SEL_BIOCHE1_RS.EOF Then
                If SEL_BIOCHE1_RS("BJPCOMMENT") <> "" Or SEL_BIOCHE1_RS("BJPCOMMENT") <> Null Then
                    txt_urrate.Text = Trim(SEL_BIOCHE1_RS("BJPCOMMENT"))
                    cmb_urrate.Text = Trim(SEL_BIOCHE1_RS("BJPCOMMENT"))
                End If
               
                If SEL_BIOCHE1_RS("KAPPAUR") <> "" Or SEL_BIOCHE1_RS("KAPPAUR") <> Null Then
                    txt_urkappa.Text = Trim(SEL_BIOCHE1_RS("KAPPAUR"))
                End If
                
                If SEL_BIOCHE1_RS("LAMBDAUR") <> "" Or SEL_BIOCHE1_RS("LAMBDAUR") <> Null Then
                    txt_urlambda.Text = Trim(SEL_BIOCHE1_RS("LAMBDAUR"))
                End If
            End If
            '********* omega 27/06/2011 *****
            
            
        If DMPROCEDUR_RS.RowCount > 0 Then
                REPORT.Enabled = False
                FRMIMM.Enabled = True 'here 1
                i = 1
                MSFlexGrid1.Rows = 1
                While Not DMPROCEDUR_RS.EOF
                    SELMA101FLAG = False
                   If SELMA101FLAG = False Then
                        SELMA101_QR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                        SELMA101FLAG = True
                    Else
                        SELMA101_QR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
                        SELMA101_RS.MoreResults
                        SELMA101_RS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        '***** Shraddha 07/04/10 *****
                        'SELECT * FROM interface.dinterface where reqno=? and TESTCODE=?
                        SEL_DINT_QR(0) = Trim(req_chglen(cmb_reqno.Text))
                        SEL_DINT_QR(1) = Trim(DMPROCEDUR_RS("PRO_CODE"))
                        Set SEL_DINT_RS = SEL_DINT_QR.OpenResultset(1, 3)
                        'SELDTEMPLETQR(1) = Trim(SEL_DINT_RS("MACHINENO"))
                        '*****************************
                    
                        str = "SELECT * FROM DISLIB.DTEMPLET1 WHERE PRO_CODE=? AND MACHINNA=?"
                        Set SELDTEMPLETQR = rdb.CreateQuery("", str)
'                        If SELDTEMPLETFLAG = False Then
                            SELDTEMPLETQR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
                            'SELDTEMPLETQR1(1) = Trim(MSFlexGrid1.TextMatrix(I, 7)) 'Uncommented by Shraddha on 07/04/10
                            SELDTEMPLETQR(1) = Trim(SEL_DINT_RS("MACHINENO"))
                            
                            If Trim(SEL_DINT_RS("MACHINENO")) = "" Then
                                MsgBox "Machine name has not been inserted in Interface"
                            End If
                            
                            Set SELDTEMPLETRS = SELDTEMPLETQR.OpenResultset(1, 3)
 '                           SELDTEMPLETFLAG = True
'                        Else
'                            SELDTEMPLETQR(0) = Trim(DMPROCEDUR_RS("PRO_CODE"))
'                            SELDTEMPLETQR(1) = Trim(SEL_DINT_RS("MACHINENO"))
'                            SELDTEMPLETRS.MoreResults
'                            SELDTEMPLETRS.Requery
                        'End If
                        If SELDTEMPLETRS.RowCount > 0 Then
                            tmflag = True
                            MSFlexGrid1.AddItem i & vbTab & Trim(SELMA101_RS(1)) & "$" & Trim(DMPROCEDUR_RS("PRO_CODE")) & vbTab & DMPROCEDUR_RS("OBSERVALUE") & vbTab & Trim(SELDTEMPLETRS("LOWVALUE")) & "-" & Trim(SELDTEMPLETRS("HIGHVALUE")) & vbTab & Trim(DMPROCEDUR_RS("pvalue")) & vbTab & Trim(DMPROCEDUR_RS("source")) & vbTab & vbTab & Trim(DMPROCEDUR_RS("REMARKS"))
                            i = i + 1
                        End If

                    End If
                DMPROCEDUR_RS.MoveNext
                Wend

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
                txt_comments.Text = " " & SELDMBCCOMMRS("COMMENTS")
            End If
      End If
'
'        '****
    '***** Shraddha 22/09/2012 *****
    SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
    Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
    If selpathdrequisitn_rs.RowCount > 0 Then
        txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
    End If
    '*******************************
          
End If
End Sub

Public Sub assign()
'On Error GoTo MSG:
Dim num As Integer
Dim sc As String

Dim TM, TM1, TM2 As String
    TM = "2.0"
    TM1 = "1.2"
    TM2 = "0.5"

rdb.BeginTrans
If flagsave = 1 Then
    flagsave = 0
    For i = 1 To MSFlexGrid1.Rows - 1
      num = InStr(Trim(MSFlexGrid1.TextMatrix(i, 1)), "$")
      sc = Mid(Trim(MSFlexGrid1.TextMatrix(i, 1)), num + 1)
        
''        If SELMA101FLAG1 = False Then
''             SELMA101_QR1(0) = Trim(MSFlexGrid1.TextMatrix(i, 1))
''             Set SELMA101_RS1 = SELMA101_QR1.OpenResultset(1, 3)
''             SELMA101FLAG1 = True
''        Else
''             SELMA101_QR1(0) = Trim(MSFlexGrid1.TextMatrix(i, 1))
''             SELMA101_RS1.MoreResults
''             SELMA101_RS1.Requery
''        End If
''       If SELMA101_RS1.RowCount > 0 Then
         If Trim(sc) <> "" Then
            
            tmflag = True
            INSDPRO(0) = req_chglen(cmb_reqno.Text) 'Insert in dmprocedur
            INSDPRO(1) = sc ' SELMA101_RS1("SER_CODE")
            
            
            
            If sc = "F824" And Trim(MSFlexGrid1.TextMatrix(i, 7)) = "ARCH" Then
                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(TM) Then
                    INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
                Else
                    INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
                End If
                
            ElseIf sc = "F819" And Trim(MSFlexGrid1.TextMatrix(i, 7)) = "ARCH" Then
                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(TM1) Then
                    INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
                Else
                    INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
                End If
                
            ElseIf sc = "F818" And Trim(MSFlexGrid1.TextMatrix(i, 7)) = "ARCH" Then
                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(TM2) Then
                    INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
                Else
                    INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
                End If
            ElseIf sc <> "F824" Or sc <> "F819" Or sc <> "F818" Then
                    INSDPRO(2) = MSFlexGrid1.TextMatrix(i, 2)
            End If
            
                If SELDTEMPLETFLAG1 = False Then
                    SELDTEMPLETQR1(0) = sc ' SELMA101_RS1("ser_code")
                    SELDTEMPLETQR1(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
                    Set SELDTEMPLETRS1 = SELDTEMPLETQR1.OpenResultset(1, 3)
                    SELDTEMPLETFLAG1 = True
                Else
                    SELDTEMPLETQR1(0) = sc ' SELMA101_RS1("ser_code")
                    SELDTEMPLETQR1(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
                    SELDTEMPLETRS1.MoreResults
                    SELDTEMPLETRS1.Requery
                End If
                If SELDTEMPLETRS1.RowCount > 0 Then
                        If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS1("HIGHVALUE"))) Then
                                INSDPRO(3) = "H"
                        
                        Else
                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS1("LOWVALUE"))) Then
                                    INSDPRO(3) = "N"
                                End If
                                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS1("LOWVALUE"))) Then
                                   INSDPRO(3) = "L"
                                End If
                                If (Trim(MSFlexGrid1.TextMatrix(i, 2))) = "" Or Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) = 0 Then
                                   INSDPRO(3) = ""
                                End If
                        End If
                        
                 INSDPRO(4) = SELDTEMPLETRS1("UOM")
                
                Else
                   INSDPRO(3) = ""
                   INSDPRO(4) = ""
                End If
            
            INSDPRO(5) = MSFlexGrid1.TextMatrix(i, 3)
            INSDPRO(6) = MSFlexGrid1.TextMatrix(i, 4)
            INSDPRO(7) = MSFlexGrid1.TextMatrix(i, 5)
            INSDPRO(8) = MSFlexGrid1.TextMatrix(i, 7)
            INSDPRO.Execute
       
           'INTERFACE FLAG=3 AND MACHINE
           
           UPDINTERFACE_QR(0) = 3
           UPDINTERFACE_QR(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
           UPDINTERFACE_QR(2) = req_chglen(cmb_reqno.Text)
           UPDINTERFACE_QR(3) = sc
           UPDINTERFACE_QR.Execute
         
         End If
    
    Next i
            '* UPDATE THE RREQUISITN
            updrrequisitn_qr(0) = Trim(cmb_enterby.Text)
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = Trim(cmb_finalby.Text)
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = Trim(cmb_finalby.Text)
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
            
            If IMFLAG = True Then
                'IMFLAG = False
             '**** inserting into imbiochelab
            INSDIMBIOCHEQR(0) = req_chglen(cmb_reqno.Text)
            INSDIMBIOCHEQR(1) = Trim(cmb_elepattern.Text)
            INSDIMBIOCHEQR(2) = Trim(cmb_monband.Text)
            INSDIMBIOCHEQR(3) = Trim(cmb_locationnband.Text)
            If Len(Trim(txt_connband.Text)) > 0 Then
                INSDIMBIOCHEQR(4) = Null
                INSDIMBIOCHEQR(4) = Trim(txt_connband.Text)
            Else
                INSDIMBIOCHEQR(4) = Null
            End If
            INSDIMBIOCHEQR(5) = txt_iga.Text
            INSDIMBIOCHEQR(6) = txt_igm.Text
            INSDIMBIOCHEQR(7) = txt_igg.Text
            If Len(Trim(txt_remmbond.Text)) > 0 Then
                INSDIMBIOCHEQR(8) = Null
                INSDIMBIOCHEQR(8) = Trim(txt_remmbond.Text)
            Else
                INSDIMBIOCHEQR(8) = Null
               
            End If
            INSDIMBIOCHEQR(9) = Trim(cmb_urineimm.Text)
            INSDIMBIOCHEQR(10) = Trim(CMB_BJPBJP.Text)
            INSDIMBIOCHEQR(11) = Trim(cmb_ka.Text)
            If Len(Trim(txt_monoclonal.Text)) > 0 Then
                INSDIMBIOCHEQR(12) = Null
                INSDIMBIOCHEQR(12) = Trim(txt_monoclonal.Text)
            Else
                INSDIMBIOCHEQR(12) = Null
                INSDIMBIOCHEQR(12) = ""
            End If
            
            If Len(Trim(txt_kappa.Text)) > 0 Then
                INSDIMBIOCHEQR(13) = Null
                INSDIMBIOCHEQR(13) = txt_kappa.Text
            Else
                INSDIMBIOCHEQR(13) = Null
                INSDIMBIOCHEQR(13) = ""
            End If
            
            If Len(Trim(txt_lambda.Text)) > 0 Then
                INSDIMBIOCHEQR(14) = Null
                INSDIMBIOCHEQR(14) = txt_lambda.Text
            Else
                INSDIMBIOCHEQR(14) = Null
                INSDIMBIOCHEQR(14) = ""
            End If
            INSDIMBIOCHEQR.Execute
                                 
            '***** Shraddha 17/03/10 *****
            upd_dinterface_qr(0) = Trim(req_chglen(cmb_reqno.Text))
            upd_dinterface_qr.Execute
            '*****************************
                                 
            '**** inserting into imbiochelab
            ''''*** Done By Kanchan
            
            INSDIMBIOCHE1(0) = req_chglen(cmb_reqno.Text)
            If Len(Trim(txt_kappafree.Text)) > 0 Then
                INSDIMBIOCHE1(1) = Trim(txt_kappafree.Text)
            Else
                INSDIMBIOCHE1(1) = Null
            End If
            
            If Len(Trim(txt_lambdafree.Text)) > 0 Then
                INSDIMBIOCHE1(2) = Trim(txt_lambdafree.Text)
            Else
                INSDIMBIOCHE1(2) = Null
            End If
            
            If Len(Trim(txt_urrate.Text)) > 0 Then
                If UCase(Trim(txt_urrate.Text)) = ".OF URINE." Then
                INSDIMBIOCHE1(3) = Null
                Else
                INSDIMBIOCHE1(3) = Trim(txt_urrate.Text)
                End If
            Else
                INSDIMBIOCHE1(3) = Null
            End If
            
            If Len(Trim(txt_urkappa.Text)) > 0 Then
                INSDIMBIOCHE1(4) = Trim(txt_urkappa.Text)
            Else
                INSDIMBIOCHE1(4) = Null
            End If
            
            If Len(Trim(txt_urlambda.Text)) > 0 Then
                INSDIMBIOCHE1(5) = Trim(txt_urlambda.Text)
            Else
                INSDIMBIOCHE1(5) = Null
            End If
            
                INSDIMBIOCHE1.Execute
            '''''''''*** Kanchan
            End If
            
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
    flagsave = 0
                 
    For i = 1 To MSFlexGrid1.Rows - 1
    num = InStr(Trim(MSFlexGrid1.TextMatrix(i, 1)), "$")
    sc = Mid(Trim(MSFlexGrid1.TextMatrix(i, 1)), num + 1)
        
''       If SELMA101FLAG1 = False Then
''            SELMA101_QR1(0) = MSFlexGrid1.TextMatrix(i, 1)
''            Set SELMA101_RS1 = SELMA101_QR1.OpenResultset(1, 3)
''            SELMA101FLAG1 = True
''       Else
''            SELMA101_QR1(0) = MSFlexGrid1.TextMatrix(i, 1)
''            SELMA101_RS1.MoreResults
''            SELMA101_RS1.Requery
''       End If
''
''       If SELMA101_RS1.RowCount > 0 Then
        If SELDTEMPLETFLAG1 = False Then
            SELDTEMPLETQR1(0) = sc 'SELMA101_RS1("ser_code")
            SELDTEMPLETQR1(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
            Set SELDTEMPLETRS1 = SELDTEMPLETQR1.OpenResultset(1, 3)
         Else
            SELDTEMPLETQR1(0) = sc 'SELMA101_RS1("ser_code")
            SELDTEMPLETRS1.MoreResults
            SELDTEMPLETRS1.Requery
            
        End If
        UPDDMPROCEDUR_QR(2) = " "  'DEFAULT UOM
        If SELDTEMPLETRS1.RowCount > 0 Then
            
            If sc = "F824" And Trim(MSFlexGrid1.TextMatrix(i, 7)) = "ARCH" Then
                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(TM) Then
                    UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                Else
                    UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                End If
                
            ElseIf sc = "F819" And Trim(MSFlexGrid1.TextMatrix(i, 7)) = "ARCH" Then
                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(TM1) Then
                    UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                Else
                    UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                End If
                
            ElseIf sc = "F818" And Trim(MSFlexGrid1.TextMatrix(i, 7)) = "ARCH" Then
                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(TM2) Then
                    UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                Else
                    UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                End If
            ElseIf sc <> "F824" Or sc <> "F819" Or sc <> "F818" Then
                    UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
            End If
            
            'UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
            If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) > Val(Trim(SELDTEMPLETRS1("HIGHVALUE"))) Then
            UPDDMPROCEDUR_QR(1) = "H"
            
            Else
                If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) >= Val(Trim(SELDTEMPLETRS1("LOWVALUE"))) Then
                    UPDDMPROCEDUR_QR(1) = "N"
                End If
                 If Val(Trim(MSFlexGrid1.TextMatrix(i, 2))) < Val(Trim(SELDTEMPLETRS1("LOWVALUE"))) Then
                    UPDDMPROCEDUR_QR(1) = "L"
                 End If
                 If (Trim(MSFlexGrid1.TextMatrix(i, 2))) = "" Then
                    UPDDMPROCEDUR_QR(1) = ""
                 End If
            End If
            
       ' End If
                    UPDDMPROCEDUR_QR(2) = Trim(SELDTEMPLETRS1("UOM"))
         End If
         
            If (Trim(MSFlexGrid1.TextMatrix(i, 2))) = "negative" Or (Trim(MSFlexGrid1.TextMatrix(i, 2))) = "positive" Then
                    UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                    UPDDMPROCEDUR_QR(1) = ""
                    UPDDMPROCEDUR_QR(2) = ""
'            Else
'                    UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
'                    UPDDMPROCEDUR_QR(1) = ""
'                    UPDDMPROCEDUR_QR(2) = ""
            End If
            
            If Trim(MSFlexGrid1.TextMatrix(i, 1)) <> "" Then 'added on 10/04/09
            '********* Added by Shraddha Pandit on 04/04/09 **********
            UPDDMPROCEDUR_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
            UPDDMPROCEDUR_QR(1) = ""
            
'            If Trim(SELDTEMPLETRS1("UOM")) <> "" Then
'            UPDDMPROCEDUR_QR(2) = Trim(SELDTEMPLETRS1("UOM")) 'added on 10/04/09
'            End If
            '*********************************************************
           ' UPDDMPROCEDUR_QR(2) = " "  'DEFAULT UOM
            UPDDMPROCEDUR_QR(3) = MSFlexGrid1.TextMatrix(i, 3)
            UPDDMPROCEDUR_QR(4) = MSFlexGrid1.TextMatrix(i, 4)
            UPDDMPROCEDUR_QR(5) = MSFlexGrid1.TextMatrix(i, 5)
            UPDDMPROCEDUR_QR(6) = MSFlexGrid1.TextMatrix(i, 7)
            UPDDMPROCEDUR_QR(7) = req_chglen(cmb_reqno.Text)
            UPDDMPROCEDUR_QR(8) = Trim(sc) 'Trim(SELMA101_RS1("SER_CODE"))
            UPDDMPROCEDUR_QR.Execute
            End If
       'End If
    Next i
    
            updrrequisitn_qr(0) = Trim(cmb_enterby.Text)
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = Trim(cmb_finalby.Text)
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = Trim(cmb_finalby.Text)
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
            
           UPDDIMBIOCHEQR(0) = Trim(cmb_elepattern.Text)
           UPDDIMBIOCHEQR(1) = Trim(cmb_monband.Text)
           UPDDIMBIOCHEQR(2) = Trim(cmb_locationnband.Text)
           If Len(Trim(txt_connband.Text)) > 0 Then
            UPDDIMBIOCHEQR(3) = Null
            UPDDIMBIOCHEQR(3) = Trim(txt_connband.Text)
           Else
            UPDDIMBIOCHEQR(3) = Null
           End If
           UPDDIMBIOCHEQR(4) = Trim(txt_iga.Text)
           UPDDIMBIOCHEQR(5) = Trim(txt_igm.Text)
           UPDDIMBIOCHEQR(6) = Trim(txt_igg.Text)
           If Len(Trim(txt_remmbond.Text)) > 0 Then
            UPDDIMBIOCHEQR(7) = Null
            UPDDIMBIOCHEQR(7) = Trim(txt_remmbond.Text)
           Else
            UPDDIMBIOCHEQR(7) = Null
            UPDDIMBIOCHEQR(7) = ""
           End If
           UPDDIMBIOCHEQR(8) = Trim(cmb_urineimm.Text)
           UPDDIMBIOCHEQR(9) = Trim(CMB_BJPBJP.Text)
           UPDDIMBIOCHEQR(10) = Trim(cmb_ka.Text)
           If Len(Trim(txt_monoclonal)) > 0 Then
            UPDDIMBIOCHEQR(11) = Null
            UPDDIMBIOCHEQR(11) = Trim(txt_monoclonal.Text)
           Else
            UPDDIMBIOCHEQR(11) = Null
            UPDDIMBIOCHEQR(11) = ""
           End If
           
           If Len(Trim(txt_kappa.Text)) > 0 Then
            UPDDIMBIOCHEQR(12) = Null
            UPDDIMBIOCHEQR(12) = Trim(txt_kappa.Text)
           Else
            UPDDIMBIOCHEQR(12) = Null
            UPDDIMBIOCHEQR(12) = ""
           End If
           
           If Len(Trim(txt_lambda.Text)) > 0 Then
            UPDDIMBIOCHEQR(13) = Null
            UPDDIMBIOCHEQR(13) = Trim(txt_lambda.Text)
           Else
            UPDDIMBIOCHEQR(13) = Null
            UPDDIMBIOCHEQR(13) = ""
           End If
           
           UPDDIMBIOCHEQR(14) = req_chglen(cmb_reqno.Text)
           UPDDIMBIOCHEQR.Execute
           
           '''*** Done By KANCHAN
           If Len(Trim(txt_kappafree.Text)) > 0 Then
                UPDDIMBIOCHE1(0) = Null
                UPDDIMBIOCHE1(0) = Trim(txt_kappafree.Text)
           Else
                UPDDIMBIOCHE1(0) = Null
                UPDDIMBIOCHE1(0) = ""
           End If
           
           If Len(Trim(txt_lambdafree.Text)) > 0 Then
                UPDDIMBIOCHE1(1) = Null
                UPDDIMBIOCHE1(1) = Trim(txt_lambdafree.Text)
           Else
                UPDDIMBIOCHE1(1) = Null
                UPDDIMBIOCHE1(1) = ""
           End If
           
           If Len(Trim(txt_urrate.Text)) > 0 Then
                UPDDIMBIOCHE1(2) = Null
                UPDDIMBIOCHE1(2) = Trim(txt_urrate.Text)
           Else
                UPDDIMBIOCHE1(2) = Null
                UPDDIMBIOCHE1(2) = ""
           End If
           
           If Len(Trim(txt_urkappa.Text)) > 0 Then
                UPDDIMBIOCHE1(3) = Null
                UPDDIMBIOCHE1(3) = Trim(txt_urkappa.Text)
           Else
                UPDDIMBIOCHE1(3) = Null
                UPDDIMBIOCHE1(3) = ""
           End If
           
           If Len(Trim(txt_urlambda.Text)) > 0 Then
                UPDDIMBIOCHE1(4) = Null
                UPDDIMBIOCHE1(4) = Trim(txt_urlambda.Text)
           Else
                UPDDIMBIOCHE1(4) = Null
                UPDDIMBIOCHE1(4) = ""
           End If
           
            UPDDIMBIOCHE1(5) = req_chglen(cmb_reqno.Text)
            UPDDIMBIOCHE1.Execute
           '''*** KANCHAN
           
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
                 UPDRTIMEM_QR(1) = txt_time.Text
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
        insertaddendum_qr(3) = txt_time.Text
        insertaddendum_qr(4) = Trim(cmb_addendumby.Text)
        insertaddendum_qr(5) = Trim(txt_addendum.Text)
        insertaddendum_qr.Execute
        
         'RTIMELOG 18-01-05
         UPDRTIMECADD_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
         UPDRTIMECADD_QR(1) = txt_time.Text
         UPDRTIMECADD_QR(2) = req_chglen(cmb_reqno.Text)
         UPDRTIMECADD_QR.Execute

   End If
   
End If

    txt_observed.Visible = False
    rdb.CommitTrans
    MsgBox "Record is  Saved Successfully."
    '''

For i = 1 To MSFlexGrid1.Rows - 1
   num = InStr(Trim(MSFlexGrid1.TextMatrix(i, 1)), "$")
   sc = Mid(Trim(MSFlexGrid1.TextMatrix(i, 1)), num + 1)
   UPDINTERFACE_QR(0) = 3
   VAR_REQNNO = req_chglen(cmb_reqno.Text)
   UPDINTERFACE_QR(1) = Trim(MSFlexGrid1.TextMatrix(i, 7))
   UPDINTERFACE_QR(2) = VAR_REQNNO
   UPDINTERFACE_QR(3) = sc
   '' COMMMENTED AS FILE IS GETTING LOCKED
   
  ' UPDINTERFACE_QR.Execute
Next i
    ''''
    btnenabl ("INI")
Exit Sub
'MSG:
Me.MousePointer = vbNormal
'rdb.RollbackTrans
'MsgBox Err.Description
End Sub

Public Sub TMADD()
    cmb_elepattern.AddItem "Normal"
    cmb_elepattern.AddItem "Abnormal"
    cmb_monband.AddItem "Present"
    cmb_monband.AddItem "Not Seen"
    cmb_locationnband.AddItem "Gamma"
    cmb_locationnband.AddItem "Beta"
    cmb_locationnband.AddItem "Alpha"
    cmb_locationnband.AddItem "Gamma+Beta"
    cmb_locationnband.AddItem "Gamma+Alpha"
    cmb_locationnband.AddItem "Alpha+beta"
    cmb_urineimm.AddItem "Normal"
    cmb_urineimm.AddItem "Ab.Normal"
    CMB_BJPBJP.AddItem "Positive"
    CMB_BJPBJP.AddItem "Negative"
    CMB_BJPBJP.AddItem "Not Requested"
    cmb_ka.AddItem "Kappa"
    cmb_ka.AddItem "Lamda"
End Sub

Public Sub assign1()
If flagsave = 1 Then
    flagsave = 0
            '**** inserting into imbiochelab
            INSDIMBIOCHEQR(0) = req_chglen(cmb_reqno.Text)
            INSDIMBIOCHEQR(1) = Trim(cmb_elepattern.Text)
            INSDIMBIOCHEQR(2) = Trim(cmb_monband.Text)
            INSDIMBIOCHEQR(3) = Trim(cmb_locationnband.Text)
            If Len(Trim(CMB_CONNBAND.Text)) > 0 Then
                INSDIMBIOCHEQR(4) = Null
                INSDIMBIOCHEQR(4) = Trim(CMB_CONNBAND.Text)
            Else
                INSDIMBIOCHEQR(4) = Null
            End If
            INSDIMBIOCHEQR(5) = txt_iga.Text
            INSDIMBIOCHEQR(6) = txt_igm.Text
            INSDIMBIOCHEQR(7) = txt_igg.Text
            If Len(Trim(txt_remmbond.Text)) > 0 Then
                INSDIMBIOCHEQR(8) = Null
                INSDIMBIOCHEQR(8) = Trim(txt_remmbond.Text)
            Else
                INSDIMBIOCHEQR(8) = Null
               
            End If
            INSDIMBIOCHEQR(9) = Trim(cmb_urineimm.Text)
            INSDIMBIOCHEQR(10) = Trim(CMB_BJPBJP.Text)
            INSDIMBIOCHEQR(11) = Trim(cmb_ka.Text)
            If Len(Trim(txt_monoclonal.Text)) > 0 Then
                INSDIMBIOCHEQR(12) = Null
                INSDIMBIOCHEQR(12) = Trim(txt_monoclonal.Text)
            Else
                INSDIMBIOCHEQR(12) = Null
                INSDIMBIOCHEQR(12) = ""
            End If
            
            If Len(Trim(txt_kappa.Text)) > 0 Then
                INSDIMBIOCHEQR(13) = Null
                INSDIMBIOCHEQR(13) = txt_kappa.Text
            Else
                INSDIMBIOCHEQR(13) = Null
                INSDIMBIOCHEQR(13) = ""
            End If
            
            If Len(Trim(txt_lambda.Text)) > 0 Then
                INSDIMBIOCHEQR(14) = Null
                INSDIMBIOCHEQR(14) = txt_lambda.Text
            Else
                INSDIMBIOCHEQR(14) = Null
                INSDIMBIOCHEQR(14) = ""
            End If
            
            'INSDIMBIOCHEQR(13) = Trim(cmb_polyclonal.Text)
            INSDIMBIOCHEQR.Execute
            '**** inserting into imbiochelab
            updrrequisitn_qr(0) = ""
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = ""
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = ""
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
            
End If
If flagsave = 2 Then
  'STR = "UPDATE DISLIB.DIMBIOCHEL SET ELEPHPATT=?,MONBAND=?,LOCNBAND=?,CONMBAND=?,IGA=?,IGM=?,IGG=?,
  'REMMBAND=?,UIMELPHO=?,BJP=?,KALA=?,TMONPIMFIX=? WHERE REQNO=?"
           UPDDIMBIOCHEQR(0) = cmb_elepattern.Text
           UPDDIMBIOCHEQR(1) = cmb_monband.Text
           UPDDIMBIOCHEQR(2) = cmb_locationnband.Text
           UPDDIMBIOCHEQR(3) = CMB_CONNBAND.Text
           UPDDIMBIOCHEQR(4) = txt_iga.Text
           UPDDIMBIOCHEQR(5) = txt_igm.Text
           UPDDIMBIOCHEQR(6) = txt_igg.Text
           If Len(Trim(txt_remmbond.Text)) > 0 Then
            UPDDIMBIOCHEQR(7) = Null
            UPDDIMBIOCHEQR(7) = txt_remmbond.Text
           Else
            UPDDIMBIOCHEQR(7) = Null
            UPDDIMBIOCHEQR(7) = ""
           End If
           UPDDIMBIOCHEQR(8) = cmb_urineimm.Text
           UPDDIMBIOCHEQR(9) = CMB_BJPBJP
           UPDDIMBIOCHEQR(10) = cmb_ka
           If Len(Trim(txt_monoclonal)) > 0 Then
            UPDDIMBIOCHEQR(11) = Null
            UPDDIMBIOCHEQR(11) = txt_monoclonal.Text
           Else
            UPDDIMBIOCHEQR(11) = Null
            UPDDIMBIOCHEQR(11) = ""
           End If
'           UPDDIMBIOCHEQR(12) = Trim(cmb_polyclonal.Text)
'           UPDDIMBIOCHEQR(13) = req_chglen(cmb_reqno.Text)
            If Len(Trim(txt_kappa.Text)) > 0 Then
            UPDDIMBIOCHEQR(12) = Null
            UPDDIMBIOCHEQR(12) = Trim(txt_kappa.Text)
           Else
            UPDDIMBIOCHEQR(12) = Null
            UPDDIMBIOCHEQR(12) = ""
           End If
           
           If Len(Trim(txt_lambda.Text)) > 0 Then
            UPDDIMBIOCHEQR(13) = Null
            UPDDIMBIOCHEQR(13) = Trim(txt_lambda.Text)
           Else
            UPDDIMBIOCHEQR(13) = Null
            UPDDIMBIOCHEQR(13) = ""
           End If
           
           UPDDIMBIOCHEQR(14) = req_chglen(cmb_reqno.Text)
           UPDDIMBIOCHEQR.Execute
           
           
            updrrequisitn_qr(0) = ""
            updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(2) = ""
            updrrequisitn_qr(3) = "T"
            updrrequisitn_qr(4) = ""
            updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
            updrrequisitn_qr(6) = Trim(req_chglen(cmb_reqno.Text))
            updrrequisitn_qr.Execute
End If
End Sub


Public Function getpreval(inuptstr As String) As String
Dim INSTR As String
Dim i As Integer
Dim j As Integer
If Len(Trim(inuptstr)) > 0 Then
   i = InStr(1, inuptstr, ",")
  If i = 0 Then
     getpreval = Trim(inuptstr)
  ' Exit Sub
  Else
  j = InStr(i + 1, inuptstr, ",")
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


