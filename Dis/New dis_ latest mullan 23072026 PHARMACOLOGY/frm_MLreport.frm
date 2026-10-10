VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_MLreport 
   BackColor       =   &H00C0C0C0&
   Caption         =   "MOLECULAR LABORATORY REPORT"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13050
   LinkTopic       =   "Form2"
   ScaleHeight     =   10170
   ScaleWidth      =   13050
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      Height          =   1215
      Left            =   120
      TabIndex        =   54
      Top             =   7800
      Width           =   10815
      Begin VB.CommandButton cmd_addendum 
         Caption         =   "Addendum"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   61
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmb_addendumby 
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
         Left            =   7560
         TabIndex        =   60
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txt_addendum 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         MultiLine       =   -1  'True
         TabIndex        =   58
         Top             =   720
         Width           =   4815
      End
      Begin VB.ComboBox cmb_finalby 
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
         Left            =   1200
         TabIndex        =   6
         Top             =   720
         Width           =   3255
      End
      Begin VB.ComboBox cmb_enterby 
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
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lbl_addendnumBy 
         BackStyle       =   0  'Transparent
         Caption         =   "Addendum by"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   59
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl_addendnum 
         BackStyle       =   0  'Transparent
         Caption         =   "Addendum"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   57
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Finalized by"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Entered  by"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Height          =   3375
      Left            =   600
      TabIndex        =   45
      Top             =   4320
      Width           =   9255
      Begin TabDlg.SSTab SSTab1 
         Height          =   2655
         Left            =   120
         TabIndex        =   63
         Top             =   720
         Visible         =   0   'False
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4683
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   -2147483645
         TabCaption(0)   =   "EXONS"
         TabPicture(0)   =   "frm_MLreport.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FRAME_EGFR"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "KRA"
         TabPicture(1)   =   "frm_MLreport.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FRAME_KRA"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame FRAME_KRA 
            BackColor       =   &H80000003&
            Height          =   2175
            Left            =   -75000
            TabIndex        =   71
            Top             =   360
            Visible         =   0   'False
            Width           =   9015
            Begin VB.ComboBox CMB_MUT_TYPE 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   6840
               TabIndex        =   75
               Top             =   360
               Width           =   1935
            End
            Begin VB.CheckBox CHK_LIT1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Caption         =   $"frm_MLreport.frx":0038
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   435
               Left            =   120
               TabIndex        =   73
               Top             =   1680
               Width           =   8775
            End
            Begin VB.ComboBox CMB_KRARESULT 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2760
               TabIndex        =   72
               Top             =   360
               Width           =   1695
            End
            Begin MSFlexGridLib.MSFlexGrid MSFLEXKRA_DATA 
               Height          =   1455
               Left            =   120
               TabIndex        =   74
               Top             =   120
               Width           =   8775
               _ExtentX        =   15478
               _ExtentY        =   2566
               _Version        =   393216
               Cols            =   5
               Appearance      =   0
               FormatString    =   $"frm_MLreport.frx":00C8
            End
         End
         Begin VB.Frame FRAME_EGFR 
            BackColor       =   &H80000003&
            Height          =   2175
            Left            =   0
            TabIndex        =   64
            Top             =   360
            Visible         =   0   'False
            Width           =   9015
            Begin VB.TextBox txt_litex12 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   8520
               MaxLength       =   2
               TabIndex        =   68
               Top             =   1680
               Width           =   375
            End
            Begin VB.TextBox txt_litex1 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   7800
               MaxLength       =   2
               TabIndex        =   67
               Top             =   1680
               Width           =   375
            End
            Begin VB.CheckBox chk_Lit 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Caption         =   "Literature suggests data supporting sensitivity to Gefitinib in activating mutations for Exon"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   315
               Left            =   120
               TabIndex        =   66
               Top             =   1680
               Width           =   7695
            End
            Begin VB.ComboBox CMB_EGFRRESULT 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   5160
               TabIndex        =   65
               Text            =   "Combo1"
               Top             =   360
               Width           =   1575
            End
            Begin MSFlexGridLib.MSFlexGrid MSFLEXEGFR_DATA 
               Height          =   1455
               Left            =   120
               TabIndex        =   69
               Top             =   120
               Width           =   8775
               _ExtentX        =   15478
               _ExtentY        =   2566
               _Version        =   393216
               Cols            =   5
               Appearance      =   0
               FormatString    =   $"frm_MLreport.frx":0157
            End
            Begin VB.Label Label8 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "and"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   375
               Left            =   8160
               TabIndex        =   70
               Top             =   1680
               Width           =   375
            End
         End
      End
      Begin VB.TextBox txt_impression 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1320
         Width           =   7455
      End
      Begin VB.ComboBox cmb_method 
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
         ItemData        =   "frm_MLreport.frx":01E8
         Left            =   1560
         List            =   "frm_MLreport.frx":01F2
         TabIndex        =   10
         Top             =   1800
         Width           =   7455
      End
      Begin VB.ComboBox cmb_result 
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
         ItemData        =   "frm_MLreport.frx":0288
         Left            =   1560
         List            =   "frm_MLreport.frx":0298
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox cmb_TestPerformed 
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
         ItemData        =   "frm_MLreport.frx":02C8
         Left            =   1560
         List            =   "frm_MLreport.frx":02CA
         TabIndex        =   2
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Method"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Impression"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Result"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Test Performed "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   240
   End
   Begin VB.Frame COMMAND 
      BackColor       =   &H80000003&
      Height          =   735
      Left            =   120
      TabIndex        =   39
      Top             =   9120
      Width           =   10815
      Begin VB.CommandButton CMD_NEW 
         Caption         =   "&New"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         Caption         =   "&Modify"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_commit 
         Caption         =   "&Commit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame common 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   600
      TabIndex        =   13
      Top             =   1200
      Width           =   9615
      Begin VB.ComboBox CMB_FZZPATHNO 
         Height          =   315
         Left            =   7440
         TabIndex        =   76
         Top             =   2520
         Width           =   1695
      End
      Begin VB.ComboBox cmb_nmaterial 
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
         ItemData        =   "frm_MLreport.frx":02CC
         Left            =   2400
         List            =   "frm_MLreport.frx":02D6
         TabIndex        =   62
         Top             =   2520
         Width           =   3975
      End
      Begin VB.TextBox txt_pathno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         TabIndex        =   52
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txt_relclnhis 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   50
         Top             =   2160
         Width           =   3975
      End
      Begin VB.ComboBox cmb_reqno 
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
         Left            =   6240
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   24
         Top             =   720
         Width           =   6615
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         TabIndex        =   22
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5880
         TabIndex        =   21
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txt_status 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txt_ward 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txt_bedno 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7920
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt_refdoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txt_procedure 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6720
         TabIndex        =   15
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox cmb_reportdoc 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   0
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2400
         TabIndex        =   0
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Path no :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   77
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Mol Lab No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   51
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nature of material"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   49
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lbl_caseno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Case No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   36
         Top             =   360
         Width           =   780
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisition No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   35
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1680
         TabIndex        =   34
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1920
         TabIndex        =   33
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3480
         TabIndex        =   32
         Top             =   1080
         Width           =   330
      End
      Begin VB.Label lbl_category 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5040
         TabIndex        =   31
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label lbl_status 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7200
         TabIndex        =   30
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label lbl_primarydoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Doctor/Service"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   29
         Top             =   1440
         Width           =   2025
      End
      Begin VB.Label lbl_ward 
         BackStyle       =   0  'Transparent
         Caption         =   "Ward"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4920
         TabIndex        =   28
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbl_bed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bed No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7200
         TabIndex        =   27
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label lbl_refdoc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referring Doctor/service"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   2115
      End
      Begin VB.Label lbl_prodis 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provisional Diagnosis"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4800
         TabIndex        =   25
         Top             =   1800
         Width           =   1815
      End
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10080
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10080
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin Crystal.CrystalReport CR 
      Left            =   0
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CRadd 
      Left            =   0
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label LBL_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOLECULAR LABORATORY MEDICAL ONCOLOGY REPORT"
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
      Left            =   840
      TabIndex        =   38
      Top             =   720
      Width           =   8820
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   960
      TabIndex        =   37
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frm_MLreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CASENO As String
 Dim ELP As Integer
Dim req_qr As rdoQuery
Dim req_rs As rdoResultset
 Dim STR_IN As String
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
 Dim INS_EGFRREPO As rdoQuery
Dim viewflag As Boolean
Dim NEWFLAG As Boolean

'*****************************************
Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset

Dim REFDR_QR As rdoQuery
Dim REFDR_RS As rdoResultset

Dim SELPATHDREQUISITN_QR As rdoQuery
Dim selpathdrequisitn_rs As rdoResultset

Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset

Dim USRNAME_QR As rdoQuery
Dim USRNAME_RS As rdoResultset

Dim LOGINUSRNAME As String
Dim LOGINTYP As String
Dim LOGINUSRID As String

Dim REPWTRNRS As rdoResultset
Dim REPWTRN_QR As rdoQuery

Dim updrrequisitn_qr As rdoQuery
Dim UPDRTIME_QR As rdoQuery

Dim INS_MOLECULAB_QR As rdoQuery

Dim sel_rtimelog_qr As rdoQuery
Dim sel_rtimelog_rs As rdoResultset

Dim flagsave As Integer

Dim CURUSRDET_QR As rdoQuery
Dim CURUSRDET_RS As rdoResultset

Dim updrrequisitn1_qr As rdoQuery
Dim UPDRTIMEC_QR As rdoQuery

Dim adenum_flg As Boolean

Dim SEL_MOLECULAB_QR As rdoQuery
Dim SEL_MOLECULAB_RS As rdoResultset

Dim UPD_MOLECULAB_QR As rdoQuery

Dim UPDRTIMECADD_QR As rdoQuery
Dim SELDADDENDUM_QR As rdoQuery
Dim SELDADDENDUM_RS As rdoResultset

Dim INSDADDENDUM_QR As rdoQuery

Dim UPD_MOLECOM_QR As rdoQuery
Dim UPD_MOLEPRN_QR As rdoQuery

Dim SEL_MA101_QR As rdoQuery
Dim SEL_MA101_RS As rdoResultset

Dim SELEGFR1_QR As rdoQuery
Dim SELEGFR1_RS As rdoResultset


'*****************************************
Sub exonMaster()
SSTab1.Visible = False
Dim SELdmulproc_QR As rdoQuery
Dim SELdmulproc_RS As rdoResultset

Dim SELEGFR2_QR As rdoQuery
Dim SELEGFR2_RS As rdoResultset


Dim TP_1 As String
TP_1 = ""
Dim I1
I1 = 0
SelStr = "SELECT * FROM DISLIB.DMULPROCE WHERE CASENO=? AND REQNO=? "
Set SELdmulproc_QR = rdb.CreateQuery("", SelStr)
SELdmulproc_QR(0) = Trim(MaskEdBox1.FormattedText)
SELdmulproc_QR(1) = Trim(Replace(cmb_reqno.Text, "/", ""))
 Set SELdmulproc_RS = SELdmulproc_QR.OpenResultset(1, 1)
   If SELdmulproc_RS.EOF = False Then
    If Trim(SELdmulproc_RS("PRO_CODE")) = "J621" Then
    SSTab1.Visible = True
    SSTab1.Tab = 0
    TP_1 = "TaqMan Probe-Based Endpoint Genotyping Mutation Analysis by Real Time PCR on the LC 480 II platform for the following Exons. Assay sensitivity is up to 2ng of DNA.Assay is run with Positive & Negative controls."
    cmb_TestPerformed.AddItem TP_1
    cmb_TestPerformed.Text = TP_1
    cmb_TestPerformed.Enabled = False
    
    
       FRAME_EGFR.Visible = True
            SelStr = "SELECT * FROM DISLIB.EGFR_MAST WHERE FLD1='E'"
                Set SELEGFR2_QR = rdb.CreateQuery("", SelStr)
              ' SELEGFR2_QR(0) = Trim(MaskEdBox1.Text)
             '   SELEGFR2_QR(1) = Trim(Replace(cmb_reqno.Text, "/", ""))
                Set SELEGFR2_RS = SELEGFR2_QR.OpenResultset(1, 1)
                If SELEGFR2_RS.EOF = False Then
                MSFLEXEGFR_DATA.clear
                MSFLEXEGFR_DATA.FormatString = "Sr.No. | EXONS         | MUTATION TYPE                                         | RESULT                   | IMPRESSION                     "
                MSFLEXEGFR_DATA.Rows = 1
                   While SELEGFR2_RS.EOF = False
                   I1 = I1 + 1
                   MSFLEXEGFR_DATA.AddItem I1 & vbTab & Trim(SELEGFR2_RS("EXONS")) & vbTab & Trim(SELEGFR2_RS("MUT_TYPE")) & vbTab & "" & vbTab & ""
                   
                   SELEGFR2_RS.MoveNext
                   
                   Wend
                   FillRes
                  
                   
                End If  'J621
     Else
     If Trim(SELdmulproc_RS("PRO_CODE")) = "J616" Or Trim(SELdmulproc_RS("PRO_CODE")) = "J618" Then
    SSTab1.Visible = True
    SSTab1.Tab = 1
    TP_1 = "TaqMan Probe-Based Endpoint Genotyping Mutation Analysis by Real Time PCR on the LC 480 II platform for the following Exons. Assay sensitivity is up to 2ng of DNA.Assay is run with Positive & Negative controls."
    cmb_TestPerformed.AddItem TP_1
    cmb_TestPerformed.Text = TP_1
    cmb_TestPerformed.Enabled = False
    
    
       FRAME_KRA.Visible = True
            SelStr = "SELECT * FROM DISLIB.EGFR_MAST WHERE FLD1='K'"
                Set SELEGFR2_QR = rdb.CreateQuery("", SelStr)
              ' SELEGFR2_QR(0) = Trim(MaskEdBox1.Text)
             '   SELEGFR2_QR(1) = Trim(Replace(cmb_reqno.Text, "/", ""))
                Set SELEGFR2_RS = SELEGFR2_QR.OpenResultset(1, 1)
                If SELEGFR2_RS.EOF = False Then
                MSFLEXKRA_DATA.clear
                MSFLEXKRA_DATA.FormatString = "Sr.No. | EXONS                             | RESULT                     | IMPRESSION                           | MUTATION TYPE             "
                MSFLEXKRA_DATA.Rows = 1
                   While SELEGFR2_RS.EOF = False
                   I1 = I1 + 1
                   MSFLEXKRA_DATA.AddItem I1 & vbTab & Trim(SELEGFR2_RS("EXONS")) & vbTab & "" & vbTab & "" & vbTab & ""
                   
                   SELEGFR2_RS.MoveNext
                   
                   Wend
                   FillRes
                  
                   
                End If  'J621
                
                
                
    End If
   End If
                       
End If
End Sub
Sub FillRes()


Dim SELEGFR3_QR As rdoQuery
Dim SELEGFR3_RS As rdoResultset

 SelStr = "SELECT DISTINCT RESULT1 FROM DISLIB.EGFR_MAST1 "
                Set SELEGFR3_QR = rdb.CreateQuery("", SelStr)
              ' SELEGFR2_QR(0) = Trim(MaskEdBox1.Text)
             '   SELEGFR2_QR(1) = Trim(Replace(cmb_reqno.Text, "/", ""))
                Set SELEGFR3_RS = SELEGFR3_QR.OpenResultset(1, 1)
                If SELEGFR3_RS.EOF = False Then
                 If SSTab1.Tab = 1 Then
                 CMB_KRARESULT.clear
                 Else
                CMB_EGFRRESULT.clear
                End If
                While SELEGFR3_RS.EOF = False
                    If SSTab1.Tab = 1 Then
                    CMB_KRARESULT.AddItem (Trim(SELEGFR3_RS(0)))
                    Else
                    CMB_EGFRRESULT.AddItem (Trim(SELEGFR3_RS(0)))
                    End If
                SELEGFR3_RS.MoveNext
                Wend
                End If
                   
End Sub

Private Sub chk_Lit_Click()
If chk_Lit.Value = 1 Then
txt_litex1.Enabled = True
txt_litex12.Enabled = True
Else
txt_litex1.Enabled = True
txt_litex12.Enabled = True
End If
End Sub

Private Sub CMB_EGFRRESULT_Click()
MSFLEXEGFR_DATA.col = 3
MSFLEXEGFR_DATA.row = MSFLEXEGFR_DATA.row
MSFLEXEGFR_DATA.Text = CMB_EGFRRESULT.Text
If CMB_EGFRRESULT.Text <> "" Then

 SelStr = "SELECT IMPRESSION FROM DISLIB.EGFR_MAST1  WHERE RESULT1=?"
   Set SELEGFR1_QR = rdb.CreateQuery("", SelStr)
  SELEGFR1_QR(0) = Trim(CMB_EGFRRESULT.Text)
'   SELEGFR2_QR(1) = Trim(Replace(cmb_reqno.Text, "/", ""))
   Set SELEGFR1_RS = SELEGFR1_QR.OpenResultset(1, 1)
   If SELEGFR1_RS.EOF = False Then
   CMB_EGFRRESULT.clear
   While SELEGFR1_RS.EOF = False
   'CMB_EGFRRESULT.AddItem (Trim(SELEGFR1_RS(0)))
     MSFLEXEGFR_DATA.col = 4
    MSFLEXEGFR_DATA.row = MSFLEXEGFR_DATA.row
    MSFLEXEGFR_DATA.Text = Trim(SELEGFR1_RS(0))

   SELEGFR1_RS.MoveNext
   Wend

   End If
End If
CMB_EGFRRESULT.Visible = False
End Sub

Private Sub CMB_KRARESULT_Change()
MSFLEXKRA_DATA.col = 2
MSFLEXKRA_DATA.row = MSFLEXKRA_DATA.row
MSFLEXKRA_DATA.Text = CMB_KRARESULT.Text
If CMB_KRARESULT.Text <> "" Then

 SelStr = "SELECT IMPRESSION FROM DISLIB.EGFR_MAST1  WHERE RESULT1=?"
   Set SELEGFR1_QR = rdb.CreateQuery("", SelStr)
  SELEGFR1_QR(0) = Trim(CMB_KRARESULT.Text)
'   SELEGFR2_QR(1) = Trim(Replace(cmb_reqno.Text, "/", ""))
   Set SELEGFR1_RS = SELEGFR1_QR.OpenResultset(1, 1)
   If SELEGFR1_RS.EOF = False Then
   
   While SELEGFR1_RS.EOF = False
   'CMB_EGFRRESULT.AddItem (Trim(SELEGFR1_RS(0)))
     MSFLEXKRA_DATA.col = 3
    MSFLEXKRA_DATA.row = MSFLEXKRA_DATA.row
    MSFLEXKRA_DATA.Text = Trim(SELEGFR1_RS(0))

   SELEGFR1_RS.MoveNext
   
            
   
   
   Wend
   
    MSFLEXKRA_DATA.col = 4
    MSFLEXKRA_DATA.row = MSFLEXKRA_DATA.row
            If Trim(MSFLEXKRA_DATA.TextMatrix(MSFLEXKRA_DATA.row, 2)) = "WILD TYPE" Then
            MSFLEXKRA_DATA.TextMatrix(MSFLEXKRA_DATA.row, 4) = "Nil"
            CMB_MUT_TYPE.Visible = False
            Else
            CMB_MUT_TYPE.Visible = True
            Dim MUT_TYPE As String
            Dim MT() As String
            SelStr = "SELECT MUT_TYPE  FROM DISLIB.EGFR_MAST  WHERE EXONS=?"
           Set SELEGFR1_QR = rdb.CreateQuery("", SelStr)
          SELEGFR1_QR(0) = Trim(MSFLEXKRA_DATA.TextMatrix(MSFLEXKRA_DATA.row, 1))
        '   SELEGFR2_QR(1) = Trim(Replace(cmb_reqno.Text, "/", ""))
           Set SELEGFR1_RS = SELEGFR1_QR.OpenResultset(1, 1)
                If SELEGFR1_RS.EOF = False Then
                On Error Resume Next
                MUT_TYPE = Trim(SELEGFR1_RS(0))
                CMB_MUT_TYPE.clear
                If InStr(MUT_TYPE, ",") Then
                MT = Split(MUT_TYPE, ",")
                    For KL = 0 To UBound(MT) - 1
                    CMB_MUT_TYPE.AddItem (MT(KL))
                    Next
                Else
                    CMB_MUT_TYPE.AddItem (MUT_TYPE)
                End If
                
                
                
                End If
            End If
    
   End If
End If
CMB_KRARESULT.Visible = False
End Sub

Private Sub CMB_KRARESULT_Click()
CMB_KRARESULT_Change
End Sub

Private Sub CMB_MUT_TYPE_Change()
MSFLEXKRA_DATA.col = 4
MSFLEXKRA_DATA.row = MSFLEXKRA_DATA.row
MSFLEXKRA_DATA.Text = CMB_MUT_TYPE.Text
CMB_MUT_TYPE.Visible = False
End Sub

Private Sub CMB_MUT_TYPE_Click()
CMB_MUT_TYPE_Change
End Sub

Private Sub cmb_reqno_Click()
If NEWFLAG = True Then
    NEWFLAG = False
        
    SELTA101_QR(0) = MaskEdBox1.FormattedText
    Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
        
    If SELTA101_RS.RowCount = 0 Then
        Exit Sub
    End If
    
    txt_sex.Text = SELTA101_RS("sex")
    txt_name.Text = Trim(SELTA101_RS("first_name")) & " " & Trim(SELTA101_RS("mid_name")) & " " & Trim(SELTA101_RS("last_name"))
    txt_age = DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
                      
    If Trim(txt_age.Text) < 1 Then
        Dim age2 As String
        age2 = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
        age = "." & (age Mod 12)
        If age2 = ".0" Then
            age2 = DateDiff("d", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
            age2 = ".0" & (age2 Mod 30)
        End If
        txt_age.Text = Trim(age2)
    End If
                
    txt_pridoc.Text = SELTA101_RS("ud_name")
    txt_category.Text = SELTA101_RS("pat_catg")
                
        
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
        
    If REFDR_RS.RowCount > 0 Then
        txt_refdoc.Text = REFDR_RS("REFBYDOC")
    End If
    
    
    SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
    Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
    
                   
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
                

    SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
    Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
                
    If selpathdrequisitn_rs.RowCount > 0 Then
        txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
        'If selpathdrequisitn_rs.EOF Then
        txt_relclnhis.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
        'End If
        txt_refdoc.Text = selpathdrequisitn_rs("REFDOC")
    End If
        

    selta1012_qr(0) = MaskEdBox1.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)

    If selta1012_rs.RowCount > 0 Then
        txt_procedure.Text = selta1012_rs("prvdiagdes")
    End If
  exonMaster
End If

        
If viewflag = True Then
    viewflag = False
        
    SELTA101_QR(0) = MaskEdBox1.FormattedText
    Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
        
    If SELTA101_RS.RowCount = 0 Then
        Exit Sub
    End If
        
    txt_sex.Text = SELTA101_RS("sex")
    txt_name.Text = Trim(SELTA101_RS("first_name")) & " " & Trim(SELTA101_RS("mid_name")) & " " & Trim(SELTA101_RS("last_name"))
    txt_age = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
    txt_pridoc.Text = SELTA101_RS("ud_name")
    txt_category.Text = SELTA101_RS("pat_catg")

    
    REFDR_QR(0) = MaskEdBox1.FormattedText
    REFDR_QR(1) = req_chglen(cmb_reqno.Text)
    Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
       
    If REFDR_RS.RowCount > 0 Then
        txt_refdoc.Text = REFDR_RS("REFBYDOC")
        If REFDR_RS("COMMITFLAG") = "Y" Then
            MsgBox "Report is commited, Can't be modified, Only Addendums can be added"
            '***** 26/03/10
            SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
            Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
            
            If Not SELDADDENDUM_RS.EOF Then
                'MsgBox "Report is commited, Can't be modified, Only Addendums can be added"
                
                txt_addendum.Text = SELDADDENDUM_RS("ADD_MAX_LN")
                cmb_addendumby.Text = SELDADDENDUM_RS("USER_ID")
                
'            ElseIf SELDADDENDUM_RS.RowCount = 0 Then
'                MsgBox "Report is commited, Can't be modified," & vbCrLf & "Only Addendum can be added"
'                adenum_flg = True
'                cmd_addendum.Visible = True

                adenum_flg = True
                cmd_addendum.Visible = True
                cmd_modify.Enabled = False
            End If
            '*****
            cmd_addendum.Visible = True
            cmd_modify.Enabled = False
            
        End If
        cmb_enterby.Text = "" & REFDR_RS("ENTERBY")
        cmb_finalby.Text = "" & REFDR_RS("FREPORTBY")
    End If
    
    '***** Read from Moleculab *****
    SEL_MOLECULAB_QR(0) = req_chglen(cmb_reqno.Text)
    Set SEL_MOLECULAB_RS = SEL_MOLECULAB_QR.OpenResultset(1, 3)
    
    If Not SEL_MOLECULAB_RS.EOF Then
        cmb_TestPerformed.Text = Trim(SEL_MOLECULAB_RS("TESTPRFORM"))
        cmb_result.Text = Trim(SEL_MOLECULAB_RS("RESULT"))
        txt_impression.Text = Trim(SEL_MOLECULAB_RS("IMPRESSION"))
        cmb_method.Text = Trim(SEL_MOLECULAB_RS("METHOD"))
        If IsNull(SEL_MOLECULAB_RS("FLD1")) = False And Trim(SEL_MOLECULAB_RS("FLD1")) <> "" Then
            cmb_nmaterial.Text = Trim(SEL_MOLECULAB_RS("FLD1")) 'Added on 16/04/10 by Shraddha
        End If
    End If
    If SEL_MOLECULAB_RS("FLD3") = Null Then
    Else
    On Error GoTo abc
    CMB_FZZPATHNO.Text = Trim(SEL_MOLECULAB_RS("FLD3"))
abc:
    End If
    
    
    SEL_MOLECULAB_RS.Close
    '*******************************
    
    SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
    Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
    
    
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
    
    
    
    SELPATHDREQUISITN_QR(0) = req_chglen(Trim(cmb_reqno.Text))
    Set selpathdrequisitn_rs = SELPATHDREQUISITN_QR.OpenResultset(1, 3)
    
    If selpathdrequisitn_rs.RowCount > 0 Then
        txt_pathno.Text = selpathdrequisitn_rs("PATHNO")
        If selpathdrequisitn_rs.EOF = False Then
            txt_relclnhis.Visible = True
            txt_relclnhis.Text = Trim(selpathdrequisitn_rs("NMETERIALS"))
        End If
    End If
    
    
    selta1012_qr(0) = MaskEdBox1.FormattedText
    Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
    
    If selta1012_rs.RowCount > 0 Then
        txt_procedure.Text = Trim(selta1012_rs("prvdiagdes"))
    End If
    Dim SELdmulproc_QR As rdoQuery
    Dim SELdmulproc_RS As rdoResultset
    SelStr = "SELECT * FROM DISLIB.DMULPROCE WHERE CASENO=? AND REQNO=? "
Set SELdmulproc_QR = rdb.CreateQuery("", SelStr)
SELdmulproc_QR(0) = Trim(MaskEdBox1.FormattedText)
SELdmulproc_QR(1) = Trim(Replace(cmb_reqno.Text, "/", ""))
 Set SELdmulproc_RS = SELdmulproc_QR.OpenResultset(1, 1)
 SSTab1.Visible = False
   If SELdmulproc_RS.EOF = False Then
    If Trim(SELdmulproc_RS("PRO_CODE")) = "J621" Then
        SSTab1.Visible = True
        SSTab1.Tab = 0
       FRAME_EGFR.Visible = True
    ElseIf Trim(SELdmulproc_RS("PRO_CODE")) = "J616" Then
            SSTab1.Visible = True
            SSTab1.Tab = 1
            FRAME_KRA.Visible = True
       
       End If
     End If
    Dim SELEGFR_REP_QR As rdoQuery
    Dim SELEGFR_REP_RS As rdoResultset
    SelStr = "SELECT * FROM DISLIB.EGFR_REPOR WHERE CASENO=? AND REQNO=? "
        Set SELEGFR_REP_QR = rdb.CreateQuery("", SelStr)
        SELEGFR_REP_QR(0) = Trim(MaskEdBox1.FormattedText)
        SELEGFR_REP_QR(1) = Trim(Replace(cmb_reqno.Text, "/", ""))
         Set SELEGFR_REP_RS = SELEGFR_REP_QR.OpenResultset(1, 1)
       If SELEGFR_REP_RS.EOF = False Then
       Dim I1 As Integer
       I1 = 0
           If Trim(SELdmulproc_RS("PRO_CODE")) = "J621" Then
           MSFLEXEGFR_DATA.clear
           MSFLEXEGFR_DATA.FormatString = "Sr.No. | EXONS         | MUTATION TYPE                                         | RESULT                   | IMPRESSION                     "
           MSFLEXEGFR_DATA.Rows = 1
           CMB_EGFRRESULT.Visible = False
                  While SELEGFR_REP_RS.EOF = False
             I1 = I1 + 1
            MSFLEXEGFR_DATA.AddItem I1 & vbTab & Trim(SELEGFR_REP_RS("EXONS")) & vbTab & Trim(SELEGFR_REP_RS("MUT_TYPE")) & vbTab & Trim(SELEGFR_REP_RS("RESULT1")) & vbTab & Trim(SELEGFR_REP_RS("IMPRESSION"))
            SELEGFR_REP_RS.MoveNext
        
            Wend
            MSFLEXEGFR_DATA.Enabled = False
        
       
       Else
             MSFLEXEGFR_DATA.clear
           MSFLEXKRA_DATA.FormatString = "Sr.No. | EXONS                             | RESULT                     | IMPRESSION                           | MUTATION TYPE             "
           MSFLEXKRA_DATA.Rows = 1
           CMB_EGFRRESULT.Visible = False
                  While SELEGFR_REP_RS.EOF = False
             I1 = I1 + 1
            MSFLEXKRA_DATA.AddItem I1 & vbTab & Trim(SELEGFR_REP_RS("EXONS")) & vbTab & Trim(SELEGFR_REP_RS("RESULT1")) & vbTab & Trim(SELEGFR_REP_RS("IMPRESSION")) & vbTab & Trim(SELEGFR_REP_RS("MUT_TYPE"))
            SELEGFR_REP_RS.MoveNext
        
            Wend
            MSFLEXKRA_DATA.Enabled = False
            
            CMB_KRARESULT.Visible = False
            CMB_MUT_TYPE.Visible = False
            
            
           End If
       
       End If
       
   End If

End Sub

Private Sub cmd_addendum_Click()
flagsave = 0
adenum_flg = True
txt_addendum.Text = ""

cmd_save.Enabled = True

FLDENABL ("ADDEN")

End Sub

Private Sub cmd_back_Click()
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_cancel_Click()
BTNENABL1 ("INI")
field_clear
'TCLEAR
viewflag = False
NEWFLAG = False
adenum_flg = False
End Sub

Public Sub field_clear()
MSFLEXKRA_DATA.clear
MSFLEXEGFR_DATA.clear
FRAME_EGFR.Visible = False
FRAME_KRA.Visible = False
SSTab1.Visible = False
cmb_reqno.clear

MaskEdBox1.Mask = ""
MaskEdBox1.Text = ""
MaskEdBox1.Mask = ">??/#####"

txt_name.Text = " "
txt_sex.Text = " "
txt_age.Text = " "
txt_category.Text = " "
txt_status.Text = " "
txt_pridoc.Text = " "
txt_ward.Text = " "
txt_bedno.Text = " "
txt_refdoc.Text = " "
txt_relclnhis.Text = " "
txt_pathno.Text = ""
txt_procedure.Text = ""

cmb_TestPerformed.Text = ""
cmb_result.Text = ""
txt_impression.Text = ""
'cmb_method.Text = ""
cmb_method.Text = ""
cmb_finalby.Text = ""
cmb_enterby.Text = ""
cmb_addendumby.Text = ""
txt_addendum.Text = ""
txt_litex1.Text = ""
txt_litex12.Text = ""
chk_Lit.Value = False
cmb_nmaterial.Text = "" 'Added on 16/04/10 by Shraddha

End Sub

Private Sub cmd_modify_Click()
flagsave = 2
adenum_flg = False
BTNENABL1 ("NEW")
FLDENABL ("MOD")

End Sub

Private Sub cmd_new_Click()
Dim L1 As Integer
'On Error GoTo MSG:
BTNENABL1 ("NEW")
FLDENABL ("NEW")
'field_clear


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

flagsave = 1
viewflag = False
NEWFLAG = True
adenum_flg = False
'FLDENABL ("ENA")

MaskEdBox1.Enabled = True
MaskEdBox1.SetFocus
'ENABL ("ENA")
'TCLEAR

Exit Sub

End Sub

Private Sub cmd_print_Click()
'PRINTDT   PRINTTM
'str = "UPDATE dislib.MOLECULAB SET PRINTDT=?,PRINTDT=? WHERE REQNO=?"

'UPD_MOLEPRN_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
'UPD_MOLEPRN_QR(1) = Format(txt_time.Text, "HH:MM:SS")
'UPD_MOLEPRN_QR(32) = req_chglen(cmb_reqno.Text)
'UPD_MOLEPRN_QR.Execute
'
'UPD_MOLECOM_QR.Close

'***** Report *****
    CR.ReportFileName = "C:\DISREPORTSECLF\MolecuReport.rpt"
    CR.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    
    CR.SelectionFormula = "{MOLECULAB.REQNO} = '" & req_chglen(cmb_reqno.Text) & "'"
    
    UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)
    
    CR.Formulas(0) = "USRID = '" & UI & "'"
    '***** 30/04/10 *****

'Commented, since its added in report formula field.
'If Trim(cmb_TestPerformed.Text) = "RT-PCR Multiplex, BCR-ABL (P190, P210)" Then
'    CR.Formulas(1) = "ReportHeading = 'Qualitative RT-PCR report'"
'ElseIf Trim(cmb_TestPerformed.Text) = "RT-PCR Nested, BCR-ABL for Follow-Up" Then
'    CR.Formulas(1) = "ReportHeading = 'Qualitative RT-PCR report'"
'ElseIf Trim(cmb_TestPerformed.Text) = "RT-PCR Multiplex, Acute Leukaemia Panel" Then
'    CR.Formulas(1) = "ReportHeading = 'Qualitative RT-PCR report'"
'ElseIf Trim(cmb_TestPerformed.Text) = "RQ-PCR PML-RARA" Then
'    CR.Formulas(1) = "ReportHeading = 'Quantitative for PML-RARA fusion Transcript'"
'ElseIf Trim(cmb_TestPerformed.Text) = "RT-PCR Nested, IGH Chain Gene Rearrangement" Then
'    CR.Formulas(1) = "ReportHeading = 'Qualitative RT-PCR report'"
'ElseIf Trim(cmb_TestPerformed.Text) = "RT-PCR Nested, TCR Gene Rearrangement" Then
'    CR.Formulas(1) = "ReportHeading = 'Qualitative RT-PCR report'"
'ElseIf Trim(cmb_TestPerformed.Text) = "RT-PCR Single Round BCR-ABL (p210)" Then
'    CR.Formulas(1) = "ReportHeading = 'Qualitative RT-PCR report'"
'ElseIf Trim(cmb_TestPerformed.Text) = "Hot Start RT-PCR for PML -RARA" Then
'    CR.Formulas(1) = "ReportHeading = 'Qualitative RT-PCR report'"
'ElseIf Trim(cmb_TestPerformed.Text) = "RT-PCR Multiplex, Acute Lymphoblastic Leukaemia Panel" Then
'    CR.Formulas(1) = "ReportHeading = 'Qualitative RT-PCR report'"
'ElseIf Trim(cmb_TestPerformed.Text) = "RT-PCR Multiplex, Acute Myeloid Leukaemia Panel" Then
'    CR.Formulas(1) = "ReportHeading = 'Qualitative RT-PCR report'"
'ElseIf Trim(cmb_TestPerformed.Text) = "RQ-PCR BCR-ABL (P210)" Then
'    CR.Formulas(1) = "ReportHeading = 'Quantitative for BCR Fusion Transcript'"
'End If
    '********************
    CR.WindowState = crptMaximized
    CR.Action = 1

''***** Addendum Report 26/03/10 *****
'    SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
'    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
'
'    If SELDADDENDUM_RS.RowCount = 1 Then
'        CRadd.ReportFileName = "C:\DISREPORTSEC\MolecuAdden.rpt"
'        CRadd.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'
'        CRadd.SelectionFormula = "{MOLECULAB.REQNO} = '" & req_chglen(cmb_reqno.Text) & "'"
'
'        UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)
'
'        'CR.Formulas(0) = "USRID = '" & UI & "'"
'
'        CRadd.WindowState = crptMaximized
'        CRadd.Action = 1
'    'ElseIf SELDADDENDUM_RS.RowCount = 0 Then
'
'    End If
'
''***************************
End Sub

Private Sub cmd_save_Click()
'***** validations *****
If Trim(cmb_TestPerformed.Text) = "" Then
    MsgBox "Please enter Test performed"
    Exit Sub
ElseIf Trim(cmb_result.Text) = "" And FRAME_EGFR.Visible = False And FRAME_KRA.Visible = False Then
    MsgBox "Please enter result"
    Exit Sub
ElseIf Trim(txt_impression.Text) = "" And FRAME_EGFR.Visible = False And FRAME_KRA.Visible = Fals Then
    MsgBox "Please enter impression"
    Exit Sub
ElseIf Trim(cmb_method.Text) = "" And FRAME_EGFR.Visible = False And FRAME_KRA.Visible = Fals Then
    MsgBox "Please enter method"
    Exit Sub
End If
'***********************
If FRAME_EGFR.Visible = True Then
    If chk_Lit.Value = True Then
        If txt_litex1.Text = "" And txt_litex12.Text = "" Then
         MsgBox "Please enter Literature Exons"
         Exit Sub
        End If
    End If
End If
assign

If adenum_flg = False Then
updrrequisitn_qr(0) = cmb_enterby.Text
updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
updrrequisitn_qr(2) = cmb_finalby.Text
updrrequisitn_qr(3) = "T"
updrrequisitn_qr(4) = cmb_finalby.Text
updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
updrrequisitn_qr(6) = req_chglen(cmb_reqno.Text)
updrrequisitn_qr.Execute

'update rtimelog
UPDRTIME_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
UPDRTIME_QR(1) = Format(txt_time.Text, "HH:MM:SS")
UPDRTIME_QR(2) = req_chglen(cmb_reqno.Text)
UPDRTIME_QR.Execute
'update rtimelog

End If
rdb.CommitTrans
If adenum_flg = True Then
    BTNENABL1 ("ADDEN")
Else
    BTNENABL1 ("SAVE")
End If
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
'On Error GoTo MSG:
 
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
        'Uncomment later
'        If CURUSRDET_RS("COMMIT") <> "YES" Then
'            MsgBox "Sorry! You are not authorised to Commit the Report of this Modality"
'            cmd_commit.SetFocus
'            Exit Sub
'        End If
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
    UPDRTIMEC_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
    UPDRTIMEC_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                    
    UPDRTIMEC_QR(2) = req_chglen(cmb_reqno.Text)
    UPDRTIMEC_QR.Execute
   'rtimelog
    
End If

'***** Update Moleculab *****
'FREPORTBY  FREPORTDT  FREPORTTM
str = "UPDATE dislib.MOLECULAB SET FREPORTBY=?,FREPORTDT=?,FREPORTTM=? WHERE REQNO=?"
Set UPD_MOLECOM_QR = rdb.CreateQuery("", str)
UPD_MOLECOM_QR(0) = cmb_finalby.Text
UPD_MOLECOM_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
UPD_MOLECOM_QR(2) = Format(txt_time.Text, "HH:MM:SS")
UPD_MOLECOM_QR(3) = req_chglen(cmb_reqno.Text)
UPD_MOLECOM_QR.Execute

UPD_MOLECOM_QR.Close
'****************************

rdb.CommitTrans
'btnenabl ("INI")
BTNENABL1 ("INI")
'TXT_CLEAR
'field_clear
cmd_print.Enabled = True
Exit Sub
'MSG:
'rdb.RollbackTrans
'MsgBox Err.Description
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
    cmd_addendum.Visible = False
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
    cmd_addendum.Visible = False
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
    'If adenum_flg = False Then
    cmd_addendum.Visible = False
    'End If
    'cmd_fullreport.Enabled = False
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
    cmd_addendum.Visible = False
    adenum_flg = False
    lbl_addendnum.Visible = False
    txt_addendum.Visible = False
    txt_addendum.Enabled = False
    lbl_addendnumBy.Visible = False
    cmb_addendumby.Visible = False
    
    
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
    cmd_addendum.Visible = False
    'cmd_fullreport.Enabled = False
ElseIf cap = "print" Then
     cmd_save.Enabled = False
     cmd_back.Enabled = True
     cmd_modify.Enabled = False
     cmd_print.Enabled = False
     cmd_cancel.Enabled = True
     CMD_NEW.Enabled = False
     cmd_addendum.Visible = False
ElseIf cap = "ADDEN" Then
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
    
    If USRNAME_RS("VIEW") = "YES" Then
        cmd_view.Enabled = True
    Else
        cmd_view.Enabled = False
    End If
    
    cmd_addendum.Visible = False
    
End If
End Sub

Public Sub assign()
'On Error GoTo MSG:


rdb.BeginTrans

 
 
 
If flagsave = 1 Then 'Insert
'6 9 8 7 5
'CASENO REQNO PATNAME SEX AGE CATG
'STATUS WARD BEDNO LABNO PRIMEDOC REFDOC PROVDIAGNO NMATERIAL TESTPRFORM
'RESULT IMPRESSION METHOD USERID IPADDR REQNDATE REQNTIME SAMCOLDATE
'SAMCOLTIME SAMRECDATE SAMRECTIME PREPORTBY PREPORTDT PREPORTTM FREPORTBY
'FREPORTDT FREPORTTM PRINTDT PRINTTM RPTHEADING

    INS_MOLECULAB_QR(0) = MaskEdBox1.FormattedText 'CASENO
    INS_MOLECULAB_QR(1) = req_chglen(cmb_reqno.Text) 'REQNO
    INS_MOLECULAB_QR(2) = Trim(txt_name.Text) 'PATNAME
    INS_MOLECULAB_QR(3) = Trim(txt_sex.Text) 'SEX
    INS_MOLECULAB_QR(4) = Trim(txt_age.Text) 'AGE
    INS_MOLECULAB_QR(5) = Trim(txt_category.Text) 'CATG
    INS_MOLECULAB_QR(6) = Trim(txt_status.Text) 'STATUS
    INS_MOLECULAB_QR(7) = Trim(txt_ward.Text) 'WARD
    INS_MOLECULAB_QR(8) = Trim(txt_bedno.Text) 'BEDNO
    INS_MOLECULAB_QR(9) = Trim(txt_pathno.Text) 'LABNO
    INS_MOLECULAB_QR(10) = Trim(txt_pridoc.Text) 'PRIMEDOC
    INS_MOLECULAB_QR(11) = Trim(txt_refdoc.Text) 'REFDOC
    INS_MOLECULAB_QR(12) = Trim(txt_procedure.Text) 'PROVDIAGNO
    INS_MOLECULAB_QR(13) = Trim(txt_relclnhis.Text) 'NMATERIAL
    INS_MOLECULAB_QR(14) = Trim(cmb_TestPerformed.Text) 'TESTPRFORM
   
  
    INS_MOLECULAB_QR(15) = Trim(cmb_result.Text) 'RESULT
    INS_MOLECULAB_QR(16) = Trim(txt_impression.Text) 'IMPRESSION
    INS_MOLECULAB_QR(17) = Trim(cmb_method.Text) 'METHOD
  
    INS_MOLECULAB_QR(18) = Trim(UID) 'USERID
    INS_MOLECULAB_QR(19) = Trim(ipadd) 'IPADDR
        
    sel_rtimelog_qr(0) = req_chglen(cmb_reqno.Text)
    Set sel_rtimelog_rs = sel_rtimelog_qr.OpenResultset(1, 3)
  
        
    If Not sel_rtimelog_rs.EOF Then
        If sel_rtimelog_rs("REQNDATE") <> "" Then
            INS_MOLECULAB_QR(20) = sel_rtimelog_rs("REQNDATE") 'REQNDATE
        Else
            INS_MOLECULAB_QR(20) = Null
        End If
        If sel_rtimelog_rs("REQNTIME") <> "" Then
            INS_MOLECULAB_QR(21) = sel_rtimelog_rs("REQNTIME") 'REQNTIME
        Else
            INS_MOLECULAB_QR(21) = Null
        End If
        If sel_rtimelog_rs("SAMCOLDATE") <> "" Then
            INS_MOLECULAB_QR(22) = sel_rtimelog_rs("SAMCOLDATE") 'SAMCOLDATE
        Else
            INS_MOLECULAB_QR(22) = Null
        End If
        If sel_rtimelog_rs("SAMCOLTIME") <> "" Then
            INS_MOLECULAB_QR(23) = sel_rtimelog_rs("SAMCOLTIME") 'SAMCOLTIME
        Else
            INS_MOLECULAB_QR(23) = Null
        End If
        If sel_rtimelog_rs("SAMRECDATE") <> "" Then
            INS_MOLECULAB_QR(24) = sel_rtimelog_rs("SAMRECDATE") 'SAMRECDATE
        Else
            INS_MOLECULAB_QR(24) = Null
        End If
        If sel_rtimelog_rs("SAMRECTIME") <> "" Then
            INS_MOLECULAB_QR(25) = sel_rtimelog_rs("SAMRECTIME") 'SAMRECTIME
        Else
            INS_MOLECULAB_QR(25) = Null
        End If
        
        INS_MOLECULAB_QR(26) = Trim(cmb_finalby.Text) 'PREPORTBY
        INS_MOLECULAB_QR(27) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY")) 'PREPORTDT
        INS_MOLECULAB_QR(28) = Format(txt_time.Text, "HH:MM:SS") 'PREPORTTM
        INS_MOLECULAB_QR(29) = "" 'FREPORTBY
              
        
        'If sel_rtimelog_rs("COMREPDATE") <> "" Then
            'INS_MOLECULAB_QR(26) = sel_rtimelog_rs("COMREPDATE")
        'Else
            INS_MOLECULAB_QR(30) = Null 'FREPORTDT
        'End If
        'If sel_rtimelog_rs("COMREPTIME") <> "" Then
            'INS_MOLECULAB_QR(27) = sel_rtimelog_rs("COMREPTIME")
        'Else
            INS_MOLECULAB_QR(31) = Null 'FREPORTTM
        'End If
        INS_MOLECULAB_QR(32) = Null 'PRINTDT
        INS_MOLECULAB_QR(33) = Null 'PRINTTM
        INS_MOLECULAB_QR(34) = Trim(cmb_method.Text) 'RPTHEADING
        
        '***** Shraddha 16/04/10 *****
        INS_MOLECULAB_QR(35) = Trim(cmb_nmaterial.Text) 'FLD1
        '*****************************
        
        
        '***** ADDED BY SHWETA  02/04/14 *****
        Dim RS_DMULPRO As rdoResultset
        Set RS_DMULPRO = rdb.OpenResultset("SELECT * FROM DISLIB.DMULPROCE WHERE REQNO='" & Trim(req_chglen(cmb_reqno.Text)) & "' AND CASENO='" & Trim(MaskEdBox1.FormattedText) & "'", 1, 1)
        If Not RS_DMULPRO.EOF Then
        INS_MOLECULAB_QR(36) = Trim(RS_DMULPRO("PRO_CODE")) 'FLD1
        Else
        INS_MOLECULAB_QR(36) = ""
        End If
        
        '*****************************
        INS_MOLECULAB_QR(37) = Trim(CMB_FZZPATHNO.Text)
        INS_MOLECULAB_QR.Execute
       
    
     
        MsgBox "RECORD SAVED SUCCESSFLY"
    End If
    
    
    If FRAME_EGFR.Visible = True Then
    Dim LIT As String
    LIT = ""
    If chk_Lit.Value = 1 Then
    LIT = chk_Lit.Caption
          If txt_litex1.Text <> "" Then
         LIT = LIT & " " & txt_litex1.Text
        End If
        
         If txt_litex1.Text <> "" And txt_litex12.Text <> "" Then
         LIT = LIT & " and"
        End If
        
        If txt_litex12.Text <> "" Then
         LIT = LIT & " " & txt_litex12.Text
        End If
        
    End If
    Dim ELP As Integer
    Dim STR_IN As String
    Dim INS_EGFRREPO As rdoQuery
      
        
        For ELP = 1 To MSFLEXEGFR_DATA.Rows - 1
        STR_IN = "INSERT INTO DISLIB.EGFR_REPOR(REQNO,EXONS,MUT_TYPE,RESULT1,IMPRESSION,CASENO,FLD5) VALUES(?,?,?,?,?,?,?)"
        Set INS_EGFRREPO = rdb.CreateQuery("", STR_IN)
        INS_EGFRREPO(0) = Trim(Replace(cmb_reqno.Text, "/", ""))
        INS_EGFRREPO(1) = Trim(MSFLEXEGFR_DATA.TextMatrix(ELP, 1))
        INS_EGFRREPO(2) = Trim(MSFLEXEGFR_DATA.TextMatrix(ELP, 2))
        INS_EGFRREPO(3) = Trim(MSFLEXEGFR_DATA.TextMatrix(ELP, 3))
        INS_EGFRREPO(4) = Trim(MSFLEXEGFR_DATA.TextMatrix(ELP, 4))
        INS_EGFRREPO(5) = MaskEdBox1.FormattedText
        INS_EGFRREPO(6) = LIT
        INS_EGFRREPO.Execute
        Next
        
    End If
    
    
    If FRAME_KRA.Visible = True Then
   
    LIT = ""
    If CHK_LIT1.Value = 1 Then
    LIT = CHK_LIT1.Caption
'          If txt_litex1.Text <> "" Then
'         LIT = LIT & " " & txt_litex1.Text
'        End If
'
'         If txt_litex1.Text <> "" And txt_litex12.Text <> "" Then
'         LIT = LIT & " and"
'        End If
'
'        If txt_litex12.Text <> "" Then
'         LIT = LIT & " " & txt_litex12.Text
'        End If
        
    End If
    

   'Dim INS_EGFRREPO As rdoQuery
      
        
        For ELP = 1 To MSFLEXKRA_DATA.Rows - 1
        STR_IN = "INSERT INTO DISLIB.EGFR_REPOR(REQNO,EXONS,MUT_TYPE,RESULT1,IMPRESSION,CASENO,FLD5) VALUES(?,?,?,?,?,?,?)"
        Set INS_EGFRREPO = rdb.CreateQuery("", STR_IN)
        INS_EGFRREPO(0) = Trim(Replace(cmb_reqno.Text, "/", ""))
        INS_EGFRREPO(1) = Trim(MSFLEXKRA_DATA.TextMatrix(ELP, 1))
        INS_EGFRREPO(2) = Trim(MSFLEXKRA_DATA.TextMatrix(ELP, 4))
        INS_EGFRREPO(3) = Trim(MSFLEXKRA_DATA.TextMatrix(ELP, 2))
        INS_EGFRREPO(4) = Trim(MSFLEXKRA_DATA.TextMatrix(ELP, 3))
        INS_EGFRREPO(5) = MaskEdBox1.FormattedText
        INS_EGFRREPO(6) = LIT
        INS_EGFRREPO.Execute
        Next
        
    End If
    
    
    
End If

If flagsave = 2 Then 'Update

'UPDATE dislib.MOLECULAB SET TESTPRFORM=?,RESULT=?,IMPRESSION=?,METHOD=? WHERE REQNO=?"
'UPDATE dislib.MOLECULAB SET TESTPRFORM=?,RESULT=?,IMPRESSION=?,METHOD=?,FLD1=? WHERE REQNO=?"

    UPD_MOLECULAB_QR(0) = Trim(cmb_TestPerformed.Text)
    UPD_MOLECULAB_QR(1) = Trim(cmb_result.Text)
    UPD_MOLECULAB_QR(2) = Trim(txt_impression.Text)
    UPD_MOLECULAB_QR(3) = Trim(cmb_method.Text)
    UPD_MOLECULAB_QR(4) = Trim(cmb_nmaterial.Text)
   
        '***** ADDED BY SHWETA  02/04/14 *****
        Dim RS_UDMULPRO As rdoResultset
        Set RS_UDMULPRO = rdb.OpenResultset("SELECT * FROM DISLIB.DMULPROCE WHERE REQNO='" & Trim(req_chglen(cmb_reqno.Text)) & "' AND CASENO='" & Trim(MaskEdBox1.FormattedText) & "'", 1, 1)
        If Not RS_UDMULPRO.EOF Then
        UPD_MOLECULAB_QR(5) = Trim(RS_UDMULPRO("PRO_CODE")) 'FLD1
        Else
        UPD_MOLECULAB_QR(5) = ""
        End If
        
        '*****************************
         UPD_MOLECULAB_QR(6) = Trim(req_chglen(cmb_reqno.Text))
    UPD_MOLECULAB_QR.Execute


  
   
   If FRAME_EGFR.Visible = True Then
    
    LIT = ""
    If chk_Lit.Value = 1 Then
    LIT = chk_Lit.Caption
         If txt_litex1.Text <> "" Then
         LIT = LIT & " " & txt_litex1.Text
        End If
        
         If txt_litex1.Text <> "" And txt_litex12.Text <> "" Then
         LIT = LIT & " and"
        End If
        
        If txt_litex12.Text <> "" Then
         LIT = LIT & " " & txt_litex12.Text
        End If
        
    End If
   
   
       If CHK_LIT1.Value = 1 Then
           LIT = CHK_LIT1.Caption
        End If
  
  
  Dim STR_DEL As String
    STR_DEL = "DELETE FROM DISLIB.EGFR_REPOR WHERE REQNO='" & Trim(Replace(cmb_reqno.Text, "/", "")) & "'"
    rdb.Execute STR_DEL
   
    For ELP = 1 To MSFLEXEGFR_DATA.Rows - 1
    STR_IN = "INSERT INTO DISLIB.EGFR_REPOR(REQNO,EXONS,MUT_TYPE,RESULT1,IMPRESSION,CASENO,fld5) VALUES(?,?,?,?,?,?,?)"
    Set INS_EGFRREPO = rdb.CreateQuery("", STR_IN)
    INS_EGFRREPO(0) = Trim(Replace(cmb_reqno.Text, "/", ""))
    INS_EGFRREPO(1) = Trim(MSFLEXEGFR_DATA.TextMatrix(ELP, 1))
    INS_EGFRREPO(2) = Trim(MSFLEXEGFR_DATA.TextMatrix(ELP, 2))
    INS_EGFRREPO(3) = Trim(MSFLEXEGFR_DATA.TextMatrix(ELP, 3))
    INS_EGFRREPO(4) = Trim(MSFLEXEGFR_DATA.TextMatrix(ELP, 4))
    INS_EGFRREPO(5) = MaskEdBox1.FormattedText
     INS_EGFRREPO(6) = LIT
    INS_EGFRREPO.Execute
    Next
    End If
    
  
       If FRAME_KRA.Visible = True Then
   
   
    
    LIT = ""
    If CHK_LIT1.Value = 1 Then
    LIT = CHK_LIT1.Caption

        
    End If
   
     Dim STR_DEL1 As String
    STR_DEL1 = "DELETE FROM DISLIB.EGFR_REPOR WHERE REQNO='" & Trim(Replace(cmb_reqno.Text, "/", "")) & "'"
    rdb.Execute STR_DEL1
   
    For ELP = 1 To MSFLEXKRA_DATA.Rows - 1
    STR_IN = "INSERT INTO DISLIB.EGFR_REPOR(REQNO,EXONS,MUT_TYPE,RESULT1,IMPRESSION,CASENO,fld5) VALUES(?,?,?,?,?,?,?)"
    Set INS_EGFRREPO = rdb.CreateQuery("", STR_IN)
    INS_EGFRREPO(0) = Trim(Replace(cmb_reqno.Text, "/", ""))
    INS_EGFRREPO(1) = Trim(MSFLEXKRA_DATA.TextMatrix(ELP, 1))
    INS_EGFRREPO(2) = Trim(MSFLEXKRA_DATA.TextMatrix(ELP, 2))
    INS_EGFRREPO(3) = Trim(MSFLEXKRA_DATA.TextMatrix(ELP, 3))
    INS_EGFRREPO(4) = Trim(MSFLEXKRA_DATA.TextMatrix(ELP, 4))
    INS_EGFRREPO(5) = MaskEdBox1.FormattedText
     INS_EGFRREPO(6) = LIT
    INS_EGFRREPO.Execute
        Next
        
      
        
    End If
   
    


    MsgBox "Record updated successfully"

End If

'***** Addendum *****
If adenum_flg = True Then
    UPDRTIMECADD_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    UPDRTIMECADD_QR(1) = time1
    UPDRTIMECADD_QR(2) = req_chglen(cmb_reqno.Text)
    UPDRTIMECADD_QR.Execute
    
    SELDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
    Set SELDADDENDUM_RS = SELDADDENDUM_QR.OpenResultset(1, 3)
    
    If SELDADDENDUM_RS.RowCount > 0 Then
        SELDADDENDUM_RS.MoveLast
        INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
        INSDADDENDUM_QR(1) = SELDADDENDUM_RS("ADDEND_NO") + 1
        INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        INSDADDENDUM_QR(3) = time1
        INSDADDENDUM_QR(4) = Trim(cmb_addendumby.Text)
        INSDADDENDUM_QR(5) = Null
        INSDADDENDUM_QR(5) = Trim(txt_addendum.Text)
        INSDADDENDUM_QR.Execute
    Else
        INSDADDENDUM_QR(0) = req_chglen(cmb_reqno.Text)
        INSDADDENDUM_QR(1) = "1"
        INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        INSDADDENDUM_QR(3) = time1
        INSDADDENDUM_QR(4) = Trim(cmb_addendumby.Text)
        INSDADDENDUM_QR(5) = Null
        INSDADDENDUM_QR(5) = Trim(txt_addendum.Text)
        INSDADDENDUM_QR.Execute
    End If
    
    MsgBox "Addendum added successfully"
    
End If
'***************************************
'MSG:
'    rdb.RollbackTrans
'    MsgBox Err.Description
End Sub

Private Sub cmd_view_Click()
'On Error GoTo MSG:
BTNENABL1 ("VIEW")

FLDENABL ("VIEW")
field_clear
'TXT_CLEAR 'Uncomment later
If DATETIMEFLAG = False Then
   datetime
   DATETIMEFLAG = True
Else
    datetime
End If
txt_date.Text = DATE11
txt_time.Text = time1


adenum_flg = False
viewflag = True
NEWFLAG = False

MaskEdBox1.Enabled = True
cmb_reqno.Enabled = True
cmd_print.Enabled = True

Exit Sub
'MSG:
'MsgBox Err.Description
End Sub

Public Sub FLDENABL(EN As String)
'If EN = "DIS" Then
'        cmb_dept.Enabled = False
'        cmb_modality.Enabled = False
'        MaskEdBox1.Enabled = False
'        cmb_reqno.Enabled = False
'        txt_pathno.Enabled = False
'        txt_name.Enabled = False
'        txt_sex.Enabled = False
'        txt_age.Enabled = False
'        txt_category.Enabled = False
'        txt_status.Enabled = False
'        txt_pridoc.Enabled = False
'        txt_ward.Enabled = False
'        txt_bedno.Enabled = False
'        txt_modcode.Enabled = False
'        txt_refdoc.Enabled = False
'        txt_relclnhis.Visible = False
'        txt_observed.Visible = False
'        txt_gttobsvd.Visible = False
'
'        cmb_machine.Visible = False
'        TXT_COMMENTS.Enabled = False
'        cmb_enterby.Enabled = False
'        cmb_finalby.Enabled = False
'End If
'If EN = "ENA" Then
'
'        MaskEdBox1.Enabled = True
'        cmb_reqno.Enabled = True
'        txt_pathno.Enabled = False
'        txt_name.Enabled = False
'        txt_sex.Enabled = False
'        txt_age.Enabled = False
'        txt_category.Enabled = False
'        txt_status.Enabled = False
'        txt_pridoc.Enabled = False
'        txt_ward.Enabled = False
'        txt_bedno.Enabled = False
'        txt_modcode.Enabled = False
'        txt_refdoc.Enabled = False
'        txt_relclnhis.Visible = True
'        txt_observed.Visible = False
'        txt_gttobsvd.Visible = False
'
'        cmb_machine.Visible = False
'        TXT_COMMENTS.Enabled = True
'        cmb_enterby.Enabled = True
'        cmb_finalby.Enabled = True
'End If
If EN = "NEW" Then
    cmb_TestPerformed.Enabled = True
    cmb_result.Enabled = True
    txt_impression.Enabled = True
    cmb_method.Enabled = True
    cmb_enterby.Enabled = False
    
    lbl_addendnum.Visible = False
    lbl_addendnumBy.Visible = False
    txt_addendum.Visible = False
    txt_addendum.Enabled = False
    cmb_addendumby.Visible = False
    cmb_addendumby.Enabled = False
    MSFLEXEGFR_DATA.Enabled = True
End If
If EN = "MOD" Then
    cmb_TestPerformed.Enabled = True
    cmb_result.Enabled = True
    txt_impression.Enabled = True
    cmb_method.Enabled = True
    cmb_enterby.Enabled = False
    
    lbl_addendnum.Visible = False
    lbl_addendnumBy.Visible = False
    txt_addendum.Visible = False
    txt_addendum.Enabled = False
    cmb_addendumby.Visible = False
    cmb_addendumby.Enabled = False
    
    cmb_nmaterial.Enabled = True
    If SSTab1.Tab = 0 Then
        MSFLEXEGFR_DATA.Enabled = True
    ElseIf SSTab1.Tab = 1 Then
    MSFLEXEGFR_DATA.Enabled = True
        
    End If
        
End If
If EN = "VIEW" Then
    cmb_TestPerformed.Enabled = False
    cmb_result.Enabled = False
    txt_impression.Enabled = False
    cmb_method.Enabled = False
    
    lbl_addendnum.Visible = False
    lbl_addendnumBy.Visible = False
    txt_addendum.Visible = False
    txt_addendum.Enabled = False
    cmb_addendumby.Visible = False
    cmb_addendumby.Enabled = False
    
    cmb_nmaterial.Enabled = False 'Added on 16/04/10 by Shraddha
    MSFLEXEGFR_DATA.Enabled = False
End If

If EN = "ADDEN" Then
    lbl_addendnum.Visible = True
    lbl_addendnumBy.Visible = True
    txt_addendum.Visible = True
    txt_addendum.Enabled = True
    cmb_addendumby.Visible = True
    cmb_addendumby.Enabled = True
    
    cmb_TestPerformed.Enabled = False
    cmb_result.Enabled = False
    txt_impression.Enabled = False
    cmb_method.Enabled = False
    MSFLEXEGFR_DATA.Enabled = False
End If

If EN = "INI" Then
    cmb_TestPerformed.Enabled = False
    cmb_result.Enabled = False
    txt_impression.Enabled = False
    cmb_method.Enabled = False

    lbl_addendnum.Visible = False
    lbl_addendnumBy.Visible = False
    txt_addendum.Visible = False
    txt_addendum.Enabled = False
    cmb_addendumby.Visible = False
    cmb_addendumby.Enabled = False
End If
End Sub


Private Sub Form_Load()

datetime
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1

cmb_method.Text = "SAMPLE ASSAY RUN WITH THE POSITIVE AND NEGATIVE CONTROLS"

viewflag = False
NEWFLAG = False
adenum_flg = False
lbl_addendnum.Visible = False
lbl_addendnumBy.Visible = False
cmd_addendum.Visible = False
cmb_addendumby.Visible = False
txt_addendum.Visible = False

'********* SECURITY 22/03/10 *****************
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
    cmd_cancel_Click
    Exit Sub
End If


str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE='CCH' AND USRTYPE>=8 ORDER BY USERNAME"
Set REPWTRN_QR = rdb.CreateQuery("", str)
Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)

If REPWTRNRS.RowCount > 0 Then
    cmb_finalby.clear
    REPWTRNRS.MoveFirst
    While Not REPWTRNRS.EOF
        If REPWTRNRS("USRTYPE") > 0 Then
            cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
        REPWTRNRS.MoveNext
    Wend
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)
'*********************************************

'for retriving reqnos
str = "select * from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E' AND MODALITYCD='ML' ORDER BY REQNDT desc"
Set req_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND MODALITYCD='ML' AND DEPT_CODE='CCH' AND (REQNSTAT='T' OR REQNSTAT='F') ORDER BY REQNDT DESC"
Set SELRREQUISITN_QR = rdb.CreateQuery("", str)

'***** 17/03/10 *****
Set SELTA101_QR = rdb.CreateQuery("", "select * from PATOOLSLIB.ta101 where case_no =?")

str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
Set REFDR_QR = rdb.CreateQuery("", str)

'FOR PATH NO
str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO =? "
Set SELPATHDREQUISITN_QR = rdb.CreateQuery("", str)

str = "select * from patoolslib.ta1012 where case_no =?"
Set selta1012_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.MOLECULAB WHERE REQNO=?"
Set SEL_MOLECULAB_QR = rdb.CreateQuery("", str)

'str = "UPDATE dislib.MOLECULAB SET TESTPRFORM=?,RESULT=?,IMPRESSION=?,METHOD=? WHERE REQNO=?"
str = "UPDATE dislib.MOLECULAB SET TESTPRFORM=?,RESULT=?,IMPRESSION=?,METHOD=?,FLD1=?,FLD2=? WHERE REQNO=?"
Set UPD_MOLECULAB_QR = rdb.CreateQuery("", str)

str = "UPDATE dislib.MOLECULAB SET FREPORTBY=?,FREPORTDT=?,FREPORTTM=? WHERE REQNO=?"
Set UPD_MOLECOM_QR = rdb.CreateQuery("", str)

str = "UPDATE dislib.MOLECULAB SET PRINTDT=?,PRINTDT=? WHERE REQNO=?"
Set UPD_MOLEPRN_QR = rdb.CreateQuery("", str)

'***** Addendum related *****
str = "update rislib.rtimelog set ADPREPDATE=?,ADPREPTIME=? where reqnno=?"
Set UPDRTIMECADD_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DADDENDUM WHERE REQNO=?"
Set SELDADDENDUM_QR = rdb.CreateQuery("", str)

'str = "INSERT INTO DISLIB.ADDENDUMCY(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
str = "INSERT INTO DISLIB.DADDENDUM(REQNO,ADDEND_NO,ADDEND_DT,ADDEND_TM,USER_ID,ADD_MAX_LN) VALUES ( ?,?,?,?,?,?)"
Set INSDADDENDUM_QR = rdb.CreateQuery("", str)
'****************************

'If sel_deptrs.RowCount > 0 Then
'    sel_deptrs.MoveFirst
'    While Not sel_deptrs.EOF
'        cmb_dept.AddItem Trim(sel_deptrs("dept_name")) & "$" & Trim(sel_deptrs("dept_code"))
'        sel_deptrs.MoveNext
'        Wend
'   End If
'End If

'***** 08/04/10 *****
str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE in ('J604','J605','J606','J607','J608','J609','J610','J611','J612','J614','J621','J616')"
Set SEL_MA101_QR = rdb.CreateQuery("", str)

Set SEL_MA101_RS = SEL_MA101_QR.OpenResultset(1, 3)

If SEL_MA101_RS.RowCount > 0 Then
    SEL_MA101_RS.MoveFirst
    While Not SEL_MA101_RS.EOF
        cmb_TestPerformed.AddItem Trim(SEL_MA101_RS("SER_DET"))
        SEL_MA101_RS.MoveNext
    Wend
    cmb_TestPerformed.AddItem "Real time RT-PCR using hybridization probes on LightCycler2.(Roche)"
End If

'********************

''Later read from MA101
''cmb_TestPerformed.AddItem "RT-PCR Multiplex, BCR-ABL (P190, P210)"
''cmb_TestPerformed.AddItem "RT-PCR Nested, BCR-ABL for Follow-Up"
''cmb_TestPerformed.AddItem "RT-PCR Multiplex, Acute Leukaemia Panel"
''cmb_TestPerformed.AddItem "RQ-PCR PML-RARA"
''cmb_TestPerformed.AddItem "RT-PCR Nested, IGH Chain Gene Rearrangement"
''cmb_TestPerformed.AddItem "RT-PCR Nested, TCR Gene Rearrangement"
''
'''***** Not in MA101 *****
''cmb_TestPerformed.AddItem "RT-PCR Single Round BCR-ABL (p210)"
''cmb_TestPerformed.AddItem "Hot Start RT-PCR for PML -RARA"
''cmb_TestPerformed.AddItem "RT-PCR Multiplex, Acute Lymphoblastic Leukaemia Panel"
''cmb_TestPerformed.AddItem "RT-PCR Multiplex, Acute Myeloid Leukaemia Panel"
'''Not in MA101 till here *****
''
''cmb_TestPerformed.AddItem "RQ-PCR BCR-ABL (P210)"

'***** 23/03/10 *****
str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE REQNNO=?"
Set updrrequisitn_qr = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn1_qr1", str)

str = "UPDATE RISLIB.RTIMELOG SET SAVREPDATE=?,SAVREPTIME=? WHERE REQNNO=?"
Set UPDRTIME_QR = rdb.CreateQuery("", str)

str = "SELECT * from rislib.rtimelog where reqnno=?"
Set sel_rtimelog_qr = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
Set UPDRTIMEC_QR = rdb.CreateQuery("QWEQWEQWEc", str)




'SELECT * FROM DISLIB/MOLECULAB
'str = "INSERT INTO dislib.MOLECULAB (CASENO,REQNO,PATNAME,SEX,AGE,CATG," & _
'"STATUS,WARD,BEDNO,LABNO,PRIMEDOC,REFDOC,PROVDIAGNO,NMATERIAL,TESTPRFORM," & _
'"RESULT,IMPRESSION,METHOD,USERID,IPADDR,REQNDATE,REQNTIME,SAMCOLDATE," & _
'"SAMCOLTIME,SAMRECDATE,SAMRECTIME,PREPORTBY,PREPORTDT,PREPORTTM,FREPORTBY," & _
'"FREPORTDT,FREPORTTM,PRINTDT,PRINTTM,RPTHEADING) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"


str = "INSERT INTO dislib.MOLECULAB (CASENO,REQNO,PATNAME,SEX,AGE,CATG," & _
"STATUS,WARD,BEDNO,LABNO,PRIMEDOC,REFDOC,PROVDIAGNO,NMATERIAL,TESTPRFORM," & _
"RESULT,IMPRESSION,METHOD,USERID,IPADDR,REQNDATE,REQNTIME,SAMCOLDATE," & _
"SAMCOLTIME,SAMRECDATE,SAMRECTIME,PREPORTBY,PREPORTDT,PREPORTTM,FREPORTBY," & _
"FREPORTDT,FREPORTTM,PRINTDT,PRINTTM,RPTHEADING,FLD1,FLD2,FLD4) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
Set INS_MOLECULAB_QR = rdb.CreateQuery("", str)

'6 9 8 7 5
'CASENO REQNO PATNAME SEX AGE CATG
'STATUS WARD BEDNO LABNO PRIMEDOC REFDOC PROVDIAGNO NMATERIAL TESTPRFORM
'RESULT IMPRESSION METHOD USERID IPADDR REQNDATE REQNTIME SAMCOLDATE
'SAMCOLTIME SAMRECDATE SAMRECTIME PREPORTBY PREPORTDT PREPORTTM FREPORTBY
'FREPORTDT FREPORTTM PRINTDT PRINTTM RPTHEADING

'********************

BTNENABL1 ("INI")
FLDENABL ("INI")

End Sub



'Public Sub btnenabl(cap As String)
'If cap = "NEW" Then
'    If Not USRNAME_RS.EOF Then
'        If USRNAME_RS("SAVE") = "YES" Then
'            cmd_save.Enabled = True
'        End If
'    End If
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
'    If Not USRNAME_RS.EOF Then
'        If USRNAME_RS("MODIFY") = "YES" Then
'            cmd_modify.Enabled = True
'        Else
'            cmd_modify.Enabled = False
'        End If
'    End If
'    CMD_NEW.Enabled = False
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    cmd_view.Enabled = False
'ElseIf cap = "SAVE" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    If Not USRNAME_RS.EOF Then
'        If USRNAME_RS("NEW") = "YES" Then
'            CMD_NEW.Enabled = True
'        Else
'            CMD_NEW.Enabled = False
'        End If
'    End If
'    cmd_save.Enabled = False
'    cmd_print.Enabled = True
'    If Not USRNAME_RS.EOF Then
'        If USRNAME_RS("COMMIT") = "YES" Then
'            cmd_commit.Enabled = True
'        Else
'            cmd_commit.Enabled = False
'        End If
'    End If
'    If Not USRNAME_RS.EOF Then
'        If USRNAME_RS("VIEW") = "YES" Then
'            cmd_view.Enabled = True
'        Else
'            cmd_view.Enabled = False
'        End If
'    End If
'
'ElseIf cap = "INI" Then
'    cmd_cancel.Enabled = True
'    cmd_back.Enabled = True
'    cmd_modify.Enabled = False
'    If Not USRNAME_RS.EOF Then
'        If USRNAME_RS("NEW") = "YES" Then
'            CMD_NEW.Enabled = True
'        Else
'            CMD_NEW.Enabled = False
'        End If
'    End If
'    cmd_save.Enabled = False
'    cmd_print.Enabled = False
'    cmd_commit.Enabled = False
'    If Not USRNAME_RS.EOF Then
'        If USRNAME_RS("VIEW") = "YES" Then
'            cmd_view.Enabled = True
'        Else
'            cmd_view.Enabled = False
'        End If
'    End If
'
'ElseIf cap = "COMMIT" Then
'    cmd_save.Enabled = False
'    cmd_cancel.Enabled = False
'    cmd_back.Enabled = True
'    If Not USRNAME_RS.EOF Then
'        If USRNAME_RS("NEW") = "YES" Then
'            CMD_NEW.Enabled = True
'        Else
'            CMD_NEW.Enabled = False
'        End If
'    End If
'    cmd_modify.Enabled = False
'    cmd_print.Enabled = True
'    cmd_commit.Enabled = False
'    If Not USRNAME_RS.EOF Then
'        If USRNAME_RS("VIEW") = "YES" Then
'            cmd_view.Enabled = True
'        Else
'            cmd_view.Enabled = False
'        End If
'    End If
'
'ElseIf cap = "print" Then
'     cmd_save.Enabled = False
'     cmd_back.Enabled = True
'     cmd_modify.Enabled = False
'     cmd_print.Enabled = False
'     cmd_cancel.Enabled = True
'     CMD_NEW.Enabled = False
'End If
'
'End Sub

Private Sub MaskEdBox1_LostFocus()
CASENO = MaskEdBox1.Text
''MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3)) ''05/12/2022
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

        
If NEWFLAG = True Then
    req_qr(0) = MaskEdBox1.FormattedText
    Set req_rs = req_qr.OpenResultset(1, 3)
        
    If req_rs.RowCount > 0 Then
        'cmb_reqno.clear
        While Not req_rs.EOF
            cmb_reqno.AddItem req_chglen(req_rs("reqnno"))
            req_rs.MoveNext
        Wend
    End If
    
    getPathno
End If
        
If viewflag = True Then
    SELRREQUISITN_QR(0) = MaskEdBox1.FormattedText
    Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)

    If SELRREQUISITN_RS.RowCount > 0 Then
        cmb_reqno.clear
        While Not SELRREQUISITN_RS.EOF
            cmb_reqno.AddItem req_chglen(SELRREQUISITN_RS("REQNNO"))
            SELRREQUISITN_RS.MoveNext
        Wend
    Else
        MsgBox "No Requisition for this Modality"
        BTNENABL1 ("INI")
        Exit Sub
    End If
    
    'viewPathno
End If



End Sub

Sub getPathno()
CMB_FZZPATHNO.clear
Dim rs_pathono As rdoResultset
    Set rs_pathono = rdb.OpenResultset("SELECT pathno,reqnNo FROM patoolslib.REQNREPORT WHERE CASENO='" & CASENO & "' AND REQNNO LIKE 'FZZSP%' AND REQNSTAT IN ('F','T') ORDER BY FREPORTDT DESC  ", 1, 1)
    If Not rs_pathono.EOF Then
        If Trim(rs_pathono("PATHNO")) <> "" Then
            While rs_pathono.EOF = False
            CMB_FZZPATHNO.AddItem (Trim(rs_pathono(0)))
               
            Wend
           End If
           End If
End Sub

Private Sub MSFLEXEGFR_DATA_Click()
If MSFLEXEGFR_DATA.col = 3 Then

If Trim(MSFLEXEGFR_DATA.TextMatrix(MSFLEXEGFR_DATA.row, 3)) <> "" Then
        CMB_EGFRRESULT.Text = Trim(MSFLEXEGFR_DATA.TextMatrix(MSFLEXEGFR_DATA.row, 3))
        End If
        CMB_EGFRRESULT.Visible = True
        CMB_EGFRRESULT.SetFocus
       ' CHK_Save.Visible = True
'        CHK_Save.SetFocus
        CMB_EGFRRESULT.Left = MSFLEXEGFR_DATA.CellLeft + MSFLEXEGFR_DATA.Left
        CMB_EGFRRESULT.Top = MSFLEXEGFR_DATA.CellTop + MSFLEXEGFR_DATA.Top
    FillRes
End If
End Sub

Private Sub MSFLEXKRA_DATA_Click()
If MSFLEXKRA_DATA.col = 2 Then

If Trim(MSFLEXKRA_DATA.TextMatrix(MSFLEXKRA_DATA.row, 2)) <> "" Then
        CMB_KRARESULT.Text = Trim(MSFLEXKRA_DATA.TextMatrix(MSFLEXKRA_DATA.row, 2))
        End If
        CMB_KRARESULT.Visible = True
        CMB_KRARESULT.SetFocus
       ' CHK_Save.Visible = True
'        CHK_Save.SetFocus
        CMB_KRARESULT.Left = MSFLEXKRA_DATA.CellLeft + MSFLEXKRA_DATA.Left
        CMB_KRARESULT.Top = MSFLEXKRA_DATA.CellTop + MSFLEXKRA_DATA.Top
    FillRes
ElseIf MSFLEXKRA_DATA.col = 4 Then

If Trim(MSFLEXKRA_DATA.TextMatrix(MSFLEXKRA_DATA.row, 4)) <> "" Then
        CMB_MUT_TYPE.Text = Trim(MSFLEXKRA_DATA.TextMatrix(MSFLEXKRA_DATA.row, 4))
        End If
        CMB_MUT_TYPE.Visible = True
        CMB_MUT_TYPE.SetFocus
       ' CHK_Save.Visible = True
'        CHK_Save.SetFocus
        CMB_MUT_TYPE.Left = MSFLEXKRA_DATA.CellLeft + MSFLEXKRA_DATA.Left
        CMB_MUT_TYPE.Top = MSFLEXKRA_DATA.CellTop + MSFLEXKRA_DATA.Top
    FillRes
    
End If
End Sub

Private Sub Timer1_Timer()

datetime
txt_time.Text = time1

End Sub

Private Sub txt_litex1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
Else
KeyAscii = 0
End If
End Sub
