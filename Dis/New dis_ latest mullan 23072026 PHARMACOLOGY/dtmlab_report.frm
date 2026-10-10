VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form dtmlab_report 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "DTM REPORT"
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
      Left            =   11160
      TabIndex        =   104
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton CMD_COMM 
      Caption         =   "Comments"
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
      Left            =   7440
      TabIndex        =   101
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame REPORT 
      Height          =   2895
      Left            =   9480
      TabIndex        =   71
      Top             =   8760
      Visible         =   0   'False
      Width           =   11775
      Begin VB.TextBox txt_rh 
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
         Left            =   2520
         TabIndex        =   82
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txt_antbdyscr 
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
         Left            =   2520
         TabIndex        =   81
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txt_antglbtstddir 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         TabIndex        =   80
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txt_bldgrp 
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
         Left            =   2520
         TabIndex        =   79
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txt_antglbtstindir 
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
         Left            =   2520
         TabIndex        =   78
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txt_hbsag 
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
         Left            =   2520
         TabIndex        =   77
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txt_antihcv 
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
         Left            =   8160
         TabIndex        =   76
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txt_antihivw 
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
         Left            =   8160
         TabIndex        =   75
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txt_antihive 
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
         Left            =   8160
         TabIndex        =   74
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txt_invst 
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
         Left            =   8160
         TabIndex        =   73
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txt_preinvst 
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
         Left            =   8160
         TabIndex        =   72
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lbl_bldgrp 
         AutoSize        =   -1  'True
         Caption         =   "Blood Group"
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
         Left            =   1200
         TabIndex        =   93
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lbl_antglbtstindir 
         AutoSize        =   -1  'True
         Caption         =   "Antiglobulin Test-Indirect"
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
         TabIndex        =   92
         Top             =   1920
         Width           =   2145
      End
      Begin VB.Label lbl_rh 
         AutoSize        =   -1  'True
         Caption         =   "Rh"
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
         Left            =   2040
         TabIndex        =   91
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lbl_antbdyscr 
         AutoSize        =   -1  'True
         Caption         =   " Antibody Screen"
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
         Left            =   825
         TabIndex        =   90
         Top             =   1080
         Width           =   1470
      End
      Begin VB.Label lbl_antglbtstdir 
         AutoSize        =   -1  'True
         Caption         =   "Antiglobulin Test-Direct"
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
         TabIndex        =   89
         Top             =   1560
         Width           =   2010
      End
      Begin VB.Label lbl_hbsag 
         AutoSize        =   -1  'True
         Caption         =   "HbsAg"
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
         TabIndex        =   88
         Top             =   2280
         Width           =   570
      End
      Begin VB.Label lbl_antihivw 
         AutoSize        =   -1  'True
         Caption         =   "Anti HIV (Western Blot)"
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
         Left            =   5880
         TabIndex        =   87
         Top             =   840
         Width           =   1995
      End
      Begin VB.Label lbl_invst 
         AutoSize        =   -1  'True
         Caption         =   "Previously Investigated?"
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
         Left            =   5880
         TabIndex        =   86
         Top             =   1680
         Width           =   2100
      End
      Begin VB.Label lbl_preinvst 
         AutoSize        =   -1  'True
         Caption         =   "Previous Investigations"
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
         Left            =   6000
         TabIndex        =   85
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label lbl_antihcv 
         AutoSize        =   -1  'True
         Caption         =   "Anti HCV"
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
         Left            =   7080
         TabIndex        =   84
         Top             =   360
         Width           =   780
      End
      Begin VB.Label antihive 
         AutoSize        =   -1  'True
         Caption         =   "Anti HIV (Elisa)"
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
         Left            =   6600
         TabIndex        =   83
         Top             =   1320
         Width           =   1305
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
      TabIndex        =   46
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
      TabIndex        =   45
      Top             =   600
      Width           =   1095
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
      TabIndex        =   38
      Top             =   6840
      Width           =   12735
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
         Left            =   1440
         TabIndex        =   41
         Top             =   240
         Width           =   9375
      End
      Begin VB.ComboBox cmb_enterby 
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
         Height          =   315
         ItemData        =   "dtmlab_report.frx":0000
         Left            =   1440
         List            =   "dtmlab_report.frx":0002
         TabIndex        =   40
         Top             =   600
         Width           =   2895
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
         Left            =   8280
         TabIndex        =   39
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
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
         Left            =   240
         TabIndex        =   44
         Top             =   360
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
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lbl_finalby 
         AutoSize        =   -1  'True
         Caption         =   "Test Done  by"
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
         Left            =   6960
         TabIndex        =   42
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Frame COMMAND 
      Height          =   795
      Left            =   0
      TabIndex        =   29
      Top             =   7800
      Width           =   12735
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
         MouseIcon       =   "dtmlab_report.frx":0004
         MousePointer    =   99  'Custom
         TabIndex        =   37
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
         MouseIcon       =   "dtmlab_report.frx":030E
         MousePointer    =   99  'Custom
         TabIndex        =   36
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
         MouseIcon       =   "dtmlab_report.frx":0618
         MousePointer    =   99  'Custom
         TabIndex        =   35
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
         MouseIcon       =   "dtmlab_report.frx":0922
         MousePointer    =   99  'Custom
         TabIndex        =   34
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
         MouseIcon       =   "dtmlab_report.frx":0C2C
         MousePointer    =   99  'Custom
         TabIndex        =   33
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
         Left            =   8640
         MouseIcon       =   "dtmlab_report.frx":0F36
         MousePointer    =   99  'Custom
         TabIndex        =   32
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
         Left            =   9960
         MouseIcon       =   "dtmlab_report.frx":1240
         MousePointer    =   99  'Custom
         TabIndex        =   31
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
         MouseIcon       =   "dtmlab_report.frx":154A
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1575
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   15015
      Begin VB.ComboBox com_dept 
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
         Left            =   1200
         TabIndex        =   14
         Top             =   240
         Width           =   3135
      End
      Begin VB.ComboBox com_modality 
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
         ItemData        =   "dtmlab_report.frx":1854
         Left            =   5160
         List            =   "dtmlab_report.frx":1856
         TabIndex        =   13
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox txtcasename 
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
         Left            =   6720
         TabIndex        =   12
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Frame Frame7 
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1335
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
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   1095
         End
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
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.OptionButton OPT_CASENO 
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
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton OPT_ALL 
         Caption         =   "All"
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
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox Com_REQNO 
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
         Left            =   9120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox cmb_option 
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
         ItemData        =   "dtmlab_report.frx":1858
         Left            =   10200
         List            =   "dtmlab_report.frx":185A
         TabIndex        =   5
         Text            =   "RCS"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CheckBox chk_awaited 
         Caption         =   "Awaited"
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
         Left            =   8400
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chk_blank 
         Caption         =   "Blank"
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
         Left            =   8400
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
      Begin MSMask.MaskEdBox msk_caseno 
         Height          =   375
         Left            =   2760
         TabIndex        =   15
         Top             =   1080
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   14
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
      Begin MSMask.MaskEdBox msk_todt 
         Height          =   375
         Left            =   6840
         TabIndex        =   16
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_fromdt 
         Height          =   375
         Left            =   3600
         TabIndex        =   17
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox FTIME 
         Height          =   255
         Left            =   9600
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TTIME 
         Height          =   255
         Left            =   10680
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Department"
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
         TabIndex        =   28
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label1 
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
         Left            =   4440
         TabIndex        =   27
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label2 
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
         Left            =   5880
         TabIndex        =   26
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
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
         Left            =   2640
         TabIndex        =   25
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lbl_fromdt 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
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
         Left            =   6000
         TabIndex        =   24
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "Req No"
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
         Left            =   8400
         TabIndex        =   23
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Tab"
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
         Left            =   9240
         TabIndex        =   22
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "TIME FROM"
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
         Left            =   9600
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "TIME TO"
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
         Left            =   10800
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox txtusername 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   0
      TabIndex        =   47
      Top             =   2760
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "RCS"
      TabPicture(0)   =   "dtmlab_report.frx":185C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "HBsAg(ELISA)"
      TabPicture(1)   =   "dtmlab_report.frx":1878
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "HCV(ELISA)"
      TabPicture(2)   =   "dtmlab_report.frx":1894
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "HIV(ELISA)"
      TabPicture(3)   =   "dtmlab_report.frx":18B0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   3720
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   14655
         Begin VB.CheckBox chk_Commit 
            Caption         =   "Commit"
            Height          =   195
            Left            =   12360
            TabIndex        =   103
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox TXT_COMM 
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
            Left            =   9840
            TabIndex        =   102
            Top             =   600
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.CommandButton cmd_iat 
            Caption         =   "IAT"
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
            Left            =   6000
            TabIndex        =   69
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton cmd_dat 
            Caption         =   "DAT"
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
            Left            =   4920
            TabIndex        =   68
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton cmd_abs 
            Caption         =   "AbS"
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
            Left            =   3720
            TabIndex        =   67
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton cmd_bldgrp 
            Caption         =   "BL.GROUP"
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
            Left            =   2280
            TabIndex        =   66
            Top             =   3120
            Width           =   1215
         End
         Begin VB.ComboBox cmb_dat 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   6720
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cmb_iat 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   7680
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   480
            Width           =   1095
         End
         Begin VB.ComboBox cmb_abo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox cmb_rh 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4680
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox cmb_abs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            ItemData        =   "dtmlab_report.frx":18CC
            Left            =   5640
            List            =   "dtmlab_report.frx":18CE
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   480
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid MSFrcs 
            Height          =   2775
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   4895
            _Version        =   393216
            Rows            =   15
            Cols            =   11
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            ForeColorFixed  =   192
            AllowUserResizing=   3
            FormatString    =   $"dtmlab_report.frx":18D0
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
      End
      Begin VB.Frame Frame3 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   56
         Top             =   360
         Width           =   11415
         Begin VB.CommandButton cmd_hbsag 
            Caption         =   "HBsAg"
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
            Left            =   9720
            TabIndex        =   58
            Top             =   1440
            Width           =   1215
         End
         Begin VB.ComboBox cmb_hbsag 
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
            Left            =   5520
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   480
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid MSFhbsAg 
            Height          =   3135
            Left            =   480
            TabIndex        =   59
            Top             =   240
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   5530
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            ForeColorFixed  =   192
            AllowUserResizing=   3
            FormatString    =   "S.No |^           Requisition No.        |^         Case No.       |^            HBsAg         |^      DTM No."
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
      End
      Begin VB.Frame Frame4 
         Height          =   3615
         Left            =   -75000
         TabIndex        =   52
         Top             =   360
         Width           =   11655
         Begin VB.CommandButton cmd_anti_hcv 
            Caption         =   "Anti - HCV"
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
            Left            =   9840
            TabIndex        =   54
            Top             =   1440
            Width           =   1215
         End
         Begin VB.ComboBox cmb_anti_hcv 
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
            Left            =   5280
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   360
            Width           =   2295
         End
         Begin MSFlexGridLib.MSFlexGrid MSFhcv 
            Height          =   3375
            Left            =   720
            TabIndex        =   55
            Top             =   120
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   5953
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            ForeColorFixed  =   192
            AllowUserResizing=   3
            FormatString    =   "S.No |^            Requisition No.          |^  Case No.   |^      Anti - HCV (ELISA)   |^       DTM No.     "
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
      End
      Begin VB.Frame Frame5 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   48
         Top             =   360
         Width           =   11535
         Begin VB.CommandButton cmd_anti_hiv 
            Caption         =   "Anti-HIV"
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
            Left            =   9720
            TabIndex        =   50
            Top             =   1440
            Width           =   1455
         End
         Begin VB.ComboBox cmb_anti_hiv 
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
            Left            =   5040
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   360
            Width           =   2295
         End
         Begin MSFlexGridLib.MSFlexGrid MSFhiv 
            Height          =   3375
            Left            =   720
            TabIndex        =   51
            Top             =   120
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   5953
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            BackColor       =   16777215
            ForeColor       =   0
            ForeColorFixed  =   192
            AllowUserResizing=   3
            FormatString    =   "S.No |^           Requisition No.       |^  Case No.   |^       Anti- HIV (ELISA)     |^       DTM No.     "
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
      TabIndex        =   100
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
      TabIndex        =   99
      Top             =   600
      Width           =   420
   End
   Begin VB.Label lbl_transmeddtm 
      AutoSize        =   -1  'True
      Caption         =   "REPORT - TRANSFUSION MEDICINE DTM "
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
      Left            =   3240
      TabIndex        =   98
      Top             =   600
      Width           =   6015
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
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
      TabIndex        =   97
      Top             =   240
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label6 
      Caption         =   "Password"
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
      Left            =   240
      TabIndex        =   96
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "HOMI BHABHA CANCER HOSPITAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2760
      TabIndex        =   95
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label10 
      Caption         =   "01/05/10"
      Height          =   255
      Left            =   120
      TabIndex        =   94
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "dtmlab_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CMDNEW4VIEW As Boolean
Dim REQFLAG2 As Boolean
Dim REQ2_QRFLAG As Boolean
Dim REQ1_QRFLAG As Boolean
Dim sel_moadlityqr As rdoQuery
Dim sel_modalityrs As rdoResultset
Dim SEL_MODALITYFLAG1 As Boolean
Dim MOD_CODE As String
Dim sel_departmentflag As Boolean
Dim sel_departmentrs As rdoResultset
Dim SELRMODALITY_QR As rdoQuery
Dim sel_casno As rdoQuery
Dim req1_qr As rdoQuery
Dim req2_qr As rdoQuery
Dim req1_rs As rdoResultset
Dim req2_rs As rdoResultset
Dim req_qr As rdoQuery
Dim req_rs As rdoResultset
Dim DTMFLAG As Boolean
Dim prepath  As String
Dim VIS As String
Dim key As Integer
Dim Drequisitn_RS2 As rdoResultset
Dim Drequisitn_qr2 As rdoQuery
Dim Drequisitn_flag2 As Boolean

Dim str As String
Dim testorder As String
Dim sel_deptqr As rdoQuery
Dim sel_deptrs As rdoResultset
Dim SELDEPTRSFLAG1 As Boolean
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
'Dim SELEMPLIB_QR As rdoQuery
'Dim SELEMPLIB_RS As rdoResultset
Dim i As Integer
'Dim insdfsrep_qr As rdoQuery
'Dim insdgrossdesc_qr As rdoQuery
'Dim insdfs_qr As rdoQuery
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG1 As Boolean

Dim SELRREQUISITN_QR1 As rdoQuery
Dim SELRREQUISITN_RS1 As rdoResultset
Dim SELRREQUISITNFLAG11 As Boolean


'Dim selempbasinfo1_qr As rdoQuery
'Dim selempbasinfo1_rs As rdoResultset
Dim updrrequisitn_qr As rdoQuery
Dim updrrequisitn1_qr As rdoQuery
Dim INSDADDENDUM_QR As rdoQuery
Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim SELDREQUISITNFLAG1 As Boolean

Dim NEWFLAG As Boolean
Dim viewflag As Boolean

Dim flagsave As Integer
Dim ADDENDUMFALG As Boolean
Dim MODIFYFLAG As Boolean
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

Dim SELDMULPROCQR As rdoQuery
Dim seldmulprocrs As rdoResultset
Dim seldmulprocflag As Boolean

'Dim updrrequisitn1_qr As rdoQuery
Dim SELDREQNOQR As rdoQuery
Dim SELDREQNORS As rdoResultset
Dim SELDREQNOFLAG As Boolean

Dim w_impression As Application     'Dim w_imp As Application    ', w_rhythm As Application, w_stsegment As Application, w_twave As Application,
Dim d_impression As Document       ', d_rhythm As Document, d_stsegment As Document, d_twave As Document, d_impression As Document

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
Dim usr As String

Dim DrequisitnQR As rdoQuery
Dim DrequisitnRS As rdoResultset
Dim DrequisitnFLAG As Boolean

Dim selreQR As rdoQuery
Dim selreRS As rdoResultset
Dim selreflag As Boolean

''VARIABLE FOR DTMTAB TABLE

Dim selDTMTabQR As rdoQuery
Dim selDTMTabRS As rdoResultset
Dim selDTMTabFLAG As Boolean

Dim selDTMTabQR2 As rdoQuery
Dim selDTMTabRS2 As rdoResultset
Dim selDTMTabFLAG2 As Boolean

Dim selDTMTabQR1 As rdoQuery
Dim selDTMTabRS1 As rdoResultset
Dim selDTMTabFLAG1 As Boolean

''VARIABLE FOR FIND REQNO FROM DTMTAB TABLE

Dim finddtmQR As rdoQuery
Dim finddtmRS As rdoResultset
Dim finddtmFLAG As Boolean

''VARIABLE FOR INSERT DATA INTO DTMTAB TABLE

Dim INSDTMTABQR As rdoQuery
Dim INSDTMTABRS As rdoResultset
Dim INSDTMTABFLAG As Boolean

Dim INSDTMNOQR As rdoQuery
Dim INSDTMNORS As rdoResultset
Dim INSDTMNOFLAG As Boolean

Dim UPDREQQR As rdoQuery
Dim UPDREQRS As rdoResultset
Dim UPDREQFLAG As Boolean

Dim INSDMBCCOMMQR As rdoQuery

Dim SELREQUIQR As rdoQuery
Dim SELREQUIRS As rdoResultset
Dim SELREQUIFLAG As Boolean

''VARIABLE FOR RREQUISITN RESULTSET
Dim VARREQQR As rdoQuery
Dim VARREQRS As rdoResultset
Dim VARREQFLAG As Boolean

''variables for dtmtab update(ABO FIELD)
Dim UPDDTMQRABO As rdoQuery
Dim UPDDTMQRABOFLAG As Boolean

''variables for dtmtab update(HBSAG FIELD)
Dim UPDDTMQRHBSAG As rdoQuery
Dim UPDDTMQRHBSAGFLAG As Boolean

''variables for dtmtab update(HCV FIELD)
Dim UPDDTMQRHCV As rdoQuery
Dim UPDDTMQRHCVFLAG As Boolean

''variables for dtmtab update(HIV FIELD)
Dim UPDDTMQRHIV As rdoQuery
Dim UPDDTMQRHIVFLAG As Boolean

'' variables for DMBCCOMM update
Dim UPDMBCCOMMQR As rdoQuery

Dim DTMVIEW1ALLQR As rdoQuery
Dim DTMVIEW1ALLRS As rdoResultset
Dim DTMVIEW1ALLFLAG As Boolean

''PATOOLS
Dim SEL_detail_qr As rdoQuery
Dim SEL_detail_rs As rdoResultset
Dim SEL_detail_flag As Boolean

''SELECT STATEMENT OF DMBCCOMM TABLE
Dim SELDMBC_QR As rdoQuery
Dim SELDMBC_RS As rdoResultset
Dim SELDMBC_FLAG As Boolean

''COMMIT
Dim CUPDREQQR As rdoQuery

''select data from dtmuser table

Dim DTMUSER_QR As rdoQuery
Dim DTMUSER_RS As rdoResultset

'SELECT DTMUSER
Dim SELDTMUSER_FLAG As Boolean
Dim SELDTMUSER_RS As rdoResultset
Dim SELDTMUSER_QR As rdoQuery

''select dtmno
Dim SELDTMNO_FLAG As Boolean
Dim SELDTMNO_RS As rdoResultset
Dim SELDTMNO_QR As rdoQuery

''VARIABLE FOR BASINFO TABLE

Dim SELEMPBASINFO_QR As rdoQuery
Dim SELEMPBASINFO_RS As rdoResultset
Dim SELEMPBASINFOFLAG As Boolean

''variables for dtmNO TABLE update
Dim UPDDTMNO_QR As rdoQuery

''DMULPROCE
Dim SELDMLPROCE_QR As rdoQuery
Dim SELDMLPROCE_RS As rdoResultset
Dim SELDMLPROCE_FLAG As Boolean

Dim Drequisitn_QR As rdoQuery
Dim Drequisitn_RS As rdoResultset
Dim Drequisitn_FLAG As Boolean

Dim Drequisitn_QR_RCS As rdoQuery
Dim Drequisitn_RS_RCS As rdoResultset
Dim Drequisitn_FLAG_RCS As Boolean

Dim Drequisitn_QR_HBSAG As rdoQuery
Dim Drequisitn_RS_HBSAG As rdoResultset
Dim Drequisitn_FLAG_HBSAG As Boolean

Dim Drequisitn_QR_HCV As rdoQuery
Dim Drequisitn_RS_HCV As rdoResultset
Dim Drequisitn_FLAG_HCV As Boolean

Dim Drequisitn_QR_HIV As rdoQuery
Dim Drequisitn_RS_HIV As rdoResultset
Dim Drequisitn_FLAG_HIV As Boolean

Dim INSDTMHIV_QR As rdoQuery

Dim UPDDTMTAB_QR As rdoQuery
Dim UPDDTMTAB_RS As rdoResultset
Dim UPDDTMTAB_FLAG As Boolean

Dim SELDTMTAB_QR As rdoQuery
Dim SELDTMTAB_RS As rdoResultset
Dim SELDTMTAB_FLAG As Boolean

''VARIABLES FOR BLANK CASES

Dim Drequisitn_QR_RCS_BLANK As rdoQuery
Dim Drequisitn_RS_RCS_BLANK As rdoResultset
Dim Drequisitn_FLAG_RCS_BLANK As Boolean

Dim Drequisitn_QR_HBSAG_BLANK As rdoQuery
Dim Drequisitn_RS_HBSAG_BLANK As rdoResultset
Dim Drequisitn_FLAG_HBSAG_BLANK As Boolean

Dim Drequisitn_QR_HIV_BLANK As rdoQuery
Dim Drequisitn_RS_HIV_BLANK As rdoResultset
Dim Drequisitn_FLAG_HIV_BLANK As Boolean

Dim Drequisitn_QR_HCV_BLANK As rdoQuery
Dim Drequisitn_RS_HCV_BLANK As rdoResultset
Dim Drequisitn_FLAG_HCV_BLANK As Boolean

''******SECURITY***********
Dim REPWTRN_QR As rdoQuery
Dim LOGINUSRNAME As String
'Dim LOGINTYP As String
Dim LOGINTYP As Integer
Dim LOGINUSRID As String
Dim LOGINCOMMIT As String
Dim REPWTRNRS As rdoResultset
Dim REPWTRIRS As rdoResultset
Dim CURUSRDET_QR As rdoQuery
Dim CURUSRDET_RS As rdoResultset

Dim SELFINALBY_QR As rdoQuery
Dim SELFINALBY_RS As rdoResultset

Dim SELEDREQSTAFFD_QR As rdoQuery
Dim SELEDREQSTAFFD_RS As rdoResultset
Dim SELEDREQSTAFFD_FLAG As Boolean
''******SECURITY***********

'JUSTIN FOR PRINT
Dim PUPDDTMTAB_QR  As rdoQuery

Dim a As Integer


'added by shweta
Dim UPDREQQR_P As rdoQuery

Dim SEL_DMBCCOMM_QR As rdoQuery
Dim SEL_DMBCCOMM_RS As rdoResultset

Dim UPD_DMBCCOMM_qr As rdoQuery

Dim INS_DMBCCOMM_QR As rdoQuery




Private Sub chk_commit_Click()
If chk_Commit.Value = 1 Then
    'a=MsgBox("Do you want to commit this Report fo"
    cmd_commit_Click
    MSFrcs.TextMatrix(MSFrcs.row, MSFrcs.Cols - 1) = "Commited"
End If



'rdb.BeginTrans
'
'
'        CUPDREQQR(0) = Trim(cmb_enterby.Text)  'cmb_finalby.Text
'        CUPDREQQR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
'        CUPDREQQR(2) = "F"
'        CUPDREQQR(3) = "Y"
'        'CUPDREQQR(4) = MSFrcs.TextMatrix(ICOUNT, 1)
'        CUPDREQQR(4) = req_chglen(Trim(Com_REQNO.Text))
'        CUPDREQQR.Execute
'
' 'SHWETA UPDATE RTIME LOG 02-04-2013
'      If UPDRTIMELOG(Trim(req_chglen(Trim(Com_REQNO.Text))), "COMMIT") = True Then
'      Else
'      GoTo MSG
'      End If
'      '''TILL HERE
'MsgBox "Record is Committed Successfully."
'rdb.CommitTrans
End Sub

Private Sub cmb_abo_Change()
MSFrcs.TextMatrix(MSFrcs.row, 3) = Trim(cmb_abo.Text)
End Sub

Private Sub cmb_abo_Click()
MSFrcs.TextMatrix(MSFrcs.row, 3) = Trim(cmb_abo.Text)
End Sub

Private Sub cmb_abs_Click()
MSFrcs.TextMatrix(MSFrcs.row, 5) = cmb_abs.Text
End Sub

Private Sub cmb_anti_hcv_Change()
MSFhcv.TextMatrix(MSFhcv.row, 3) = Trim(cmb_anti_hcv.Text)
End Sub

Private Sub cmb_anti_hcv_Click()
MSFhcv.TextMatrix(MSFhcv.row, 3) = cmb_anti_hcv.Text
End Sub

Private Sub cmb_anti_hiv_Click()
MSFhiv.TextMatrix(MSFhiv.row, 3) = cmb_anti_hiv.Text
End Sub

Private Sub cmb_dat_Click()
MSFrcs.TextMatrix(MSFrcs.row, 6) = cmb_dat.Text
End Sub

Private Sub cmb_hbsag_Click()
MSFhbsAg.TextMatrix(MSFhbsAg.row, 3) = cmb_hbsag.Text
End Sub

Private Sub cmb_iat_Click()
MSFrcs.TextMatrix(MSFrcs.row, 7) = cmb_iat.Text
End Sub

Private Sub cmb_option_Click()
If Opt_patient.Value = True Or opt_staff.Value = True Then
If CMDNEW4VIEW = True Then
'On Error GoTo AA
DrequisitnQR.Close
AA:
'str = "select * from dislib.DTMVIEW where reqndt BETWEEN ? AND ? AND REQNSTAT='E' and MODALITYCD='SS' ORDER BY CASENO"
'ABOVE STATEMENT COMMENTED AND NEXT ADDED BY Shraddha on 31/08/2009 - change - DTMVIEW1
If viewflag = True Then
'following line commented and new line added by shweta 28-05-2013
 '    str = "select * from dislib.DTMVIEW1 where reqndt BETWEEN ? AND ? AND REQNSTAT='E' and MODALITYCD='SS' ORDER BY CASENO"
 str = "select * from dislib.DTMVIEW1 where preportdt BETWEEN ? AND ? AND REQNSTAT='T' and MODALITYCD='SS' ORDER BY CASENO"
ElseIf NEWFLAG = True Then
    str = "select * from dislib.DTMVIEW where reqndt BETWEEN ? AND ? AND REQNSTAT='E' and MODALITYCD='SS' ORDER BY CASENO"
End If
'Above if condition added by Shraddha on 12/02/10 NEW-DTMVIEW and VIEW-DTMVIEW1

'str = "select * from dislib.DTMVIEW where reqndt BETWEEN ? AND ? AND REQNSTAT='E' and MODALITYCD='SS' ORDER BY CASENO"
'Commented and added by Shraddha on 10/02/10 - DTMVIEW instead of DTMVIEW1
'as told by Kalpana m'am
'Shraddha 12/02/10 back to DTMVIEW1 to DTMVIEW
'Set DrequisitnQR = rdb.CreateQuery("DREQUISITN1", str)
'Commented and added by Shraddha on 10/02/10 - DREQUISITN1
Set DrequisitnQR = rdb.CreateQuery("", str)

End If
    MSFrcs.Rows = 15
    SSTab1.Enabled = True
    'TXT_COMM.Visible = True
    
    If NEWFLAG = True Then
        NEWFLAG = False
        If OPT_ALL.Value = True Then
            NEWALL 'PROCEDURE FOR SELECT ALL NEW CASES BETWEEN TWO DATES
        ElseIf OPT_CASENO.Value = True Then
            NEWCASE 'PROCEDURE FOR SELECT ALL NEW CASES FOR A PARTICULAR CASENO
        End If
    End If
    If viewflag = True Then
        viewflag = False
        i = 1
        If OPT_CASENO.Value = True Then
            If chk_awaited.Value = 1 Then
                AWAITEDCASE 'PROCEDURE FOR SELECT ALL SAVED AWAITED CASES FOR A PARTICULAR CASENO
            Else
                VIEWCASE 'PROCEDURE FOR SELECT ALL SAVED CASES FOR A PARTICULAR CASENO
            End If
        End If
        If OPT_ALL.Value = True Then
            If chk_awaited.Value = 1 Then
                AWAITEDALL 'PROCEDURE FOR SELECT ALL SAVED AWAITED CASES BETWEEN TWO DATES
            ElseIf chk_blank.Value = 1 Then
                BLANKALL
            Else
            
                If FTIME.Text = "" Or TTIME.Text = "" Then
                    MsgBox "Please enter time (24hr format)"
                    Exit Sub
                Else
                    VIEWALL 'PROCEDURE FOR SELECT ALL SAVED CASES BETWEEN TWO DATES
                End If
            End If ''chk_awaited
        End If ''OPT_ALL
    End If 'view end
End If
 '''patient condition
End Sub

Private Sub cmb_rh_Change()
MSFrcs.TextMatrix(MSFrcs.row, 4) = Trim(cmb_rh.Text)
End Sub

Private Sub cmb_rh_Click()
MSFrcs.TextMatrix(MSFrcs.row, 4) = Trim(cmb_rh.Text)
End Sub

Private Sub CMD_COMM_Click()
If MODIFYFLAG = True Then
    MSFrcs.row = 1
    MSFrcs.col = 9
    TXT_COMM.Visible = True
    'cmb_iat.ListIndex = -1
    TXT_COMM.Top = MSFrcs.CellTop + MSFrcs.Top
    TXT_COMM.Left = MSFrcs.CellLeft
    cmb_abo.Visible = False
    cmb_iat.Visible = False
    cmb_rh.Visible = False
    cmb_dat.Visible = False
    cmb_abs.Visible = False
    cmb_iat.Visible = False
    
    TXT_COMM.SetFocus
End If
End Sub

Private Sub cmd_commit_Click()
CMDNEW4VIEW = False
Dim i As Integer
Dim ICOUNT As Integer, RCOUNT As Integer
ICOUNT = 1
FLDENABL ("COMMIT")
If OPT_CASENO.Value = True And Trim(Com_REQNO.Text) <> "" Then
  If MsgBox("Do U want to Commit the Record ??", vbYesNo) = vbYes Then
    On Error GoTo MSG:
    ''********SECURITY***************************
    CURUSRDET_QR(0) = Trim(cmb_enterby.Text)  'Trim(cmb_finalby.Text)
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
                'MsgBox "Sorry! This Report is not to be Finalized by You"
                'cmd_commit.SetFocus
                'Exit Sub
            End If
        End If
        
        'If CURUSRDET_RS("COMMIT") <> "YES" Then 'COMM and add for chk commit autority to loginuser by Rahul dated on 01.01.2018
        If LOGINCOMMIT <> "YES" Then
            MsgBox "Sorry! You are not authorised to Commit the Report of this Modality"
            cmd_commit.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Sorry! The Security detials of this user is not available check up in DISLNKUSR"
        cmd_commit.SetFocus
        Exit Sub
    End If
''********SECURITY*************************
    rdb.BeginTrans
'JUSTIN 09-09-2004
' If cmb_option.Text = "RCS" Then
'   RCOUNT = MSFrcs.Row
'   While RCOUNT >= ICOUNT
'    If Len(MSFrcs.TextMatrix(ICOUNT, 3)) <> 0 And Len(MSFrcs.TextMatrix(ICOUNT, 4)) <> 0 Then
        
        CUPDREQQR(0) = LOGINUSRNAME 'Trim(cmb_enterby.Text)  'cmb_finalby.Text
        CUPDREQQR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        CUPDREQQR(2) = "F"
        CUPDREQQR(3) = "Y"
        'CUPDREQQR(4) = MSFrcs.TextMatrix(ICOUNT, 1)
        CUPDREQQR(4) = req_chglen(Trim(Com_REQNO.Text))
        CUPDREQQR.Execute
        
        
        'SHWETA UPDATE RTIME LOG 02-04-2013
      If UPDRTIMELOG(Trim(req_chglen(Trim(Com_REQNO.Text))), "COMMIT") = True Then
      Else
      GoTo MSG
      End If
      '''TILL HERE
        
        
        
        MsgBox "Record is Committed Successfully."
        rdb.CommitTrans
        
'    End If
'   ICOUNT = ICOUNT + 1
'  Wend
'        MsgBox "Record is Committed Successfully."
'        rdb.CommitTrans
' End If
'JUSTIN

End If
End If
cmd_print.Enabled = True
cmd_commit.Enabled = False
btnenabl ("SAVE")
Exit Sub
MSG:
rdb.RollbackTrans
MsgBox Err.Description
End Sub

Private Sub cmd_print_Click()
CMDNEW4VIEW = False
'On Error GoTo MSG:

Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim i As Integer
Dim VAR_REQNNO As String
Dim pcount As Integer

'On Error GoTo prtmsg:
If OPT_ALL.Value = True Then
    DoEvents
    pcount = 0
    If cmb_option.Text = "RCS" Then
        For i = 1 To MSFrcs.Rows - 1
            If Trim(MSFrcs.TextMatrix(i, 1)) <> "" Then
                pcount = pcount + 1
            End If
        Next i
        
        
        For i = 1 To pcount 'MSFrcs.Rows - 1
            SELDTMNO_QR.QueryTimeout = 150
            SELDTMTAB_QR.QueryTimeout = 300
            SELREQUIQR.QueryTimeout = 300
            SELDTMNO_QR.QueryTimeout = 300
            '***** Shraddha 10/02/10 *****
            req2_qr(0) = req_chglen(Trim(MSFrcs.TextMatrix(i, 1)))
            Set req2_rs = req2_qr.OpenResultset(1, 3)
            
            If Trim(req2_rs("reqnstat")) = "T" Or Trim(req2_rs("reqnstat")) = "F" Then
            '*****************************
            
            If SELDTMTAB_FLAG = False Then
                SELDTMTAB_QR(0) = req_chglen(Trim(MSFrcs.TextMatrix(i, 1)))
                Set SELDTMTAB_RS = SELDTMTAB_QR.OpenResultset(1, 3)
                SELDTMTAB_FLAG = True
            Else
                SELDTMTAB_QR(0) = req_chglen(Trim(MSFrcs.TextMatrix(i, 1)))
                SELDTMTAB_RS.MoreResults
                SELDTMTAB_RS.Requery
            End If
            
     'JUSTIN  02-09-04 CHECKING FOR PRINTED RECORD
     
'       If UCase(Trim(SELDTMTAB_RS("HCV"))) <> "Y" Or IsNull(Trim(SELDTMTAB_RS("HCV"))) Then
'
'            PUPDDTMTAB_QR(0) = "Y"
'            PUPDDTMTAB_QR(1) = req_chglen(Trim(MSFrcs.TextMatrix(i, 1)))
'            PUPDDTMTAB_QR.Execute
            
            
            MOD_CODE = Mid(MSFrcs.TextMatrix(i, 1), 5, 2)
            VAR_REQNNO = req_chglen(MSFrcs.TextMatrix(i, 1))
            If VARREQFLAG = False Then
                VARREQQR(0) = MSFrcs.TextMatrix(i, 2)
                VARREQQR(1) = VAR_REQNNO
                Set VARREQRS = VARREQQR.OpenResultset(1, 3)
                VARREQFLAG = True
            Else
                VARREQQR(0) = MSFrcs.TextMatrix(i, 2)
                VARREQQR(1) = VAR_REQNNO
                VARREQRS.MoreResults
                VARREQRS.Requery
            End If
            If SELREQUIFLAG = False Then
                SELREQUIQR(0) = MSFrcs.TextMatrix(i, 2)
                SELREQUIQR(1) = MOD_CODE
                'SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELREQUIQR(2) = "T"
                SELREQUIQR(3) = "F"
                Set SELREQUIRS = SELREQUIQR.OpenResultset(1, 3)
                SELREQUIFLAG = True
            Else
                SELREQUIQR(0) = MSFrcs.TextMatrix(i, 2)
                SELREQUIQR(1) = MOD_CODE
                'SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELREQUIQR(2) = "T"
                SELREQUIQR(3) = "F"
                SELREQUIRS.MoreResults
                SELREQUIRS.Requery
            End If
            If SELREQUIRS.RowCount > 0 Then
                prepath = ""
                While Not SELREQUIRS.EOF
                    If Trim(SELREQUIRS("reqno")) <> Trim(VAR_REQNNO) Then
                        prepath = prepath & req_chglen(SELREQUIRS("REQNO")) & " "
                    End If
                    SELREQUIRS.MoveNext
                Wend
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = MSFrcs.TextMatrix(i, 2)
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = MSFrcs.TextMatrix(i, 2)
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            
            'If opt_patient.Value = True Then
            If Len(Trim(MSFrcs.TextMatrix(i, 2))) = 8 Then
                str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtm.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
                End If
                CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
                If SELDTMNO_RS.EOF Then
                    CrystalReport1.Formulas(3) = "dtmno = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(3) = "dtmno = '" & SELDTMNO_RS("dtmno") & "'"
                End If
                If Not SELDTMTAB_RS.EOF Then
                    a = CInt(SELDTMTAB_RS("hbsagkey"))
                
                    If Len(Trim(SELDTMTAB_RS("HBSAG"))) > 0 Then
                        CrystalReport1.Formulas(4) = "HBSAG = '" & DECRPT(Trim(SELDTMTAB_RS("hbsag")), a) & "'"
                    Else
                        CrystalReport1.Formulas(4) = "HBSAG = '" & "" & "'"
                    End If
                    If Len(Trim(SELDTMTAB_RS("HCV"))) > 0 Then
                        CrystalReport1.Formulas(5) = "HCV = '" & SELDTMTAB_RS("HCV") & "'"
                    Else
                        CrystalReport1.Formulas(5) = "HCV = '" & "" & "'"
                    End If
                    If Len(Trim(SELDTMTAB_RS("HIV"))) > 0 Then
                        CrystalReport1.Formulas(6) = "HIV = '" & SELDTMTAB_RS("HIV") & "'"
                    Else
                        CrystalReport1.Formulas(6) = "HIV = '" & "" & "'"
                    End If
                End If
            End If
            
'**** STAFF
            'If opt_staff.Value = True Then
            
            If Len(Trim(MSFrcs.TextMatrix(i, 2))) = 6 Then
                If SELEMPBASINFOFLAG = False Then
                    SELEMPBASINFO_QR(0) = MSFrcs.TextMatrix(i, 2)
                    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
                    SELEMPBASINFOFLAG = True
                Else
                    SELEMPBASINFO_QR(0) = MSFrcs.TextMatrix(i, 2)
                    SELEMPBASINFO_RS.MoreResults
                    SELEMPBASINFO_RS.Requery
                End If
                If REQ2_QRFLAG = False Then
                    req2_qr(0) = req_chglen(MSFrcs.TextMatrix(i, 1))
                    Set req2_rs = req2_qr.OpenResultset(1, 3)
                    REQ2_QRFLAG = True
                Else
                    req2_qr(0) = req_chglen(MSFrcs.TextMatrix(i, 1))
                    req2_rs.MoreResults
                    req2_rs.Requery
                End If
                str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(MSFrcs.TextMatrix(i, 1)) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtmstaff.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & MSFrcs.TextMatrix(i, 1) & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
                End If
                CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(3) = "NAM= '" & Trim(txtcasename.Text) & "'"
                CrystalReport1.Formulas(4) = "DOB= '" & DateDiff("YYYY", SELEMPBASINFO_RS("dob"), req2_rs("FREPORTDT")) & "'"
                CrystalReport1.Formulas(5) = "STDOB= '" & SELEMPBASINFO_RS("gender") & "'"
                a = CInt(SELDTMTAB_RS("hbsagkey"))
                If Len(Trim(SELDTMTAB_RS("HBSAG"))) > 0 Then
                    CrystalReport1.Formulas(6) = "HBSAG = '" & DECRPT(Trim(SELDTMTAB_RS("hbsag")), a) & "'"
                Else
                    CrystalReport1.Formulas(6) = "HBSAG = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HCVKEY"))
                If Len(Trim(SELDTMTAB_RS("HCV"))) > 0 Then
                    CrystalReport1.Formulas(7) = "HCV = '" & DECRPT(Trim(SELDTMTAB_RS("HCV")), a) & "'"
                Else
                    CrystalReport1.Formulas(7) = "HCV = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HIVKEY"))
                If Len(Trim(SELDTMTAB_RS("HIV"))) > 0 Then
                    CrystalReport1.Formulas(8) = "HIV = '" & DECRPT(Trim(SELDTMTAB_RS("HIV")), a) & "'"
                Else
                    CrystalReport1.Formulas(8) = "HIV = '" & "" & "'"
                End If
            End If
'***** STAFF
            'CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Destination = crptToPrinter
            'CrystalReport1.Destination = crptToWindow
            If Not SELDTMTAB_RS.EOF Then
                If Trim(SELDTMTAB_RS("ABO")) <> "" Or Trim(SELDTMTAB_RS("RH")) <> "" Or Trim(SELDTMTAB_RS("ABS")) <> "" Or Trim(SELDTMTAB_RS("DAT")) <> "" Or Trim(SELDTMTAB_RS("IAT")) <> "" Or Trim(SELDTMTAB_RS("HBSAG")) <> "" Or Trim(SELDTMTAB_RS("HCV")) <> "" Or Trim(SELDTMTAB_RS("HIV")) <> "" Then
            End If
               'CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
          'JUSTIN 02-09-2004 END CHECK
          'End If
          
          End If '10/02/10
        Next i
    End If
    If cmb_option.Text = "HBsAg(ELISA)" Then
        For i = 1 To MSFhbsAg.Rows - 1
            SELDTMNO_QR.QueryTimeout = 150
            SELDTMTAB_QR.QueryTimeout = 300
            SELREQUIQR.QueryTimeout = 300
            SELDTMNO_QR.QueryTimeout = 300
            If Trim(MSFhbsAg.TextMatrix(i, 1)) <> "" Then
                pcount = pcount + 1
            End If
        Next i
        For i = 1 To pcount 'MSFhbsAg.Rows - 1
            VAR_REQNNO = req_chglen(MSFhbsAg.TextMatrix(i, 1))
            If SELDTMTAB_FLAG = False Then
                SELDTMTAB_QR(0) = VAR_REQNNO
                Set SELDTMTAB_RS = SELDTMTAB_QR.OpenResultset(1, 3)
                SELDTMTAB_FLAG = True
            Else
                SELDTMTAB_QR(0) = VAR_REQNNO
                SELDTMTAB_RS.MoreResults
                SELDTMTAB_RS.Requery
            End If
            
       'JUSTIN  02-09-04 CHECKING FOR PRINTED RECORD
     
'       If UCase(Trim(SELDTMTAB_RS("HCV"))) <> "Y" Or IsNull(Trim(SELDTMTAB_RS("HCV"))) Then
'
'            PUPDDTMTAB_QR(0) = "Y"
'            PUPDDTMTAB_QR(1) = VAR_REQNNO
'            PUPDDTMTAB_QR.Execute
            
            
            MOD_CODE = Mid(MSFhbsAg.TextMatrix(i, 1), 5, 2)
            
            If VARREQFLAG = False Then
                VARREQQR(0) = MSFhbsAg.TextMatrix(i, 2)
                VARREQQR(1) = VAR_REQNNO
                Set VARREQRS = VARREQQR.OpenResultset(1, 3)
                VARREQFLAG = True
            Else
                VARREQQR(0) = MSFhbsAg.TextMatrix(i, 2)
                VARREQQR(1) = VAR_REQNNO
                VARREQRS.MoreResults
                VARREQRS.Requery
            End If
            If SELREQUIFLAG = False Then
                SELREQUIQR(0) = MSFhbsAg.TextMatrix(i, 2)
                SELREQUIQR(1) = MOD_CODE
                'SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELREQUIQR(2) = "T"
                SELREQUIQR(3) = "F"
                Set SELREQUIRS = SELREQUIQR.OpenResultset(1, 3)
                SELREQUIFLAG = True
            Else
                SELREQUIQR(0) = MSFhbsAg.TextMatrix(i, 2)
                SELREQUIQR(1) = MOD_CODE
                'SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELREQUIQR(2) = "T"
                SELREQUIQR(3) = "F"
                SELREQUIRS.MoreResults
                SELREQUIRS.Requery
            End If
            If SELREQUIRS.RowCount > 0 Then
                prepath = ""
                While Not SELREQUIRS.EOF
                    If Trim(SELREQUIRS("reqno")) <> Trim(VAR_REQNNO) Then
                        prepath = prepath & req_chglen(SELREQUIRS("REQNO")) & " "
                    End If
                    SELREQUIRS.MoveNext
                Wend
            End If
            'Dim pcount As Integer
            pcount = 0
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = MSFhbsAg.TextMatrix(i, 2)
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = MSFhbsAg.TextMatrix(i, 2)
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If Opt_patient.Value = True Then
                str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtm.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
                End If
                CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
                If SELDTMNO_RS.EOF Then
                    CrystalReport1.Formulas(3) = "dtmno = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(3) = "dtmno = '" & SELDTMNO_RS("dtmno") & "'"
                End If
                a = CInt(SELDTMTAB_RS("hbsagkey"))
                If Len(Trim(SELDTMTAB_RS("HBSAG"))) > 0 Then
                    CrystalReport1.Formulas(4) = "HBSAG = '" & DECRPT(Trim(SELDTMTAB_RS("hbsag")), a) & "'"
                Else
                    CrystalReport1.Formulas(4) = "HBSAG = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HCVKEY"))
                If Len(Trim(SELDTMTAB_RS("HCV"))) > 0 Then
                    CrystalReport1.Formulas(5) = "HCV = '" & DECRPT(Trim(SELDTMTAB_RS("HCV")), a) & "'"
                Else
                    CrystalReport1.Formulas(5) = "HCV = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HIVKEY"))
                If Len(Trim(SELDTMTAB_RS("HIV"))) > 0 Then
                    CrystalReport1.Formulas(6) = "HIV = '" & DECRPT(Trim(SELDTMTAB_RS("HIV")), a) & "'"
                Else
                    CrystalReport1.Formulas(6) = "HIV = '" & "" & "'"
                End If
            End If
'**** STAFF
            If opt_staff.Value = True Then
                If SELEMPBASINFOFLAG = False Then
                    SELEMPBASINFO_QR(0) = MSFhbsAg.TextMatrix(i, 2)
                    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
                    SELEMPBASINFOFLAG = True
                Else
                    SELEMPBASINFO_QR(0) = MSFhbsAg.TextMatrix(i, 2)
                    SELEMPBASINFO_RS.MoreResults
                    SELEMPBASINFO_RS.Requery
                End If
                If REQ2_QRFLAG = False Then
                    req2_qr(0) = req_chglen(MSFhbsAg.TextMatrix(i, 1))
                    Set req2_rs = req2_qr.OpenResultset(1, 3)
                    REQ2_QRFLAG = True
                Else
                    req2_qr(0) = req_chglen(MSFhbsAg.TextMatrix(i, 1))
                    req2_rs.MoreResults
                    req2_rs.Requery
                End If
                str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(MSFhbsAg.TextMatrix(i, 1)) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtmstaff.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & MSFhbsAg.TextMatrix(i, 1) & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
                End If
                CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(3) = "NAM= '" & Trim(txtcasename.Text) & "'"
                CrystalReport1.Formulas(4) = "DOB= '" & DateDiff("YYYY", SELEMPBASINFO_RS("dob"), req2_rs("FREPORTDT")) & "'"
                CrystalReport1.Formulas(5) = "STDOB= '" & SELEMPBASINFO_RS("gender") & "'"
                a = CInt(SELDTMTAB_RS("hbsagkey"))
                If Len(Trim(SELDTMTAB_RS("HBSAG"))) > 0 Then
                    CrystalReport1.Formulas(6) = "HBSAG = '" & DECRPT(Trim(SELDTMTAB_RS("hbsag")), a) & "'"
                Else
                    CrystalReport1.Formulas(6) = "HBSAG = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HCVKEY"))
                If Len(Trim(SELDTMTAB_RS("HCV"))) > 0 Then
                    CrystalReport1.Formulas(7) = "HCV = '" & DECRPT(Trim(SELDTMTAB_RS("HCV")), a) & "'"
                Else
                    CrystalReport1.Formulas(7) = "HCV = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HIVKEY"))
                If Len(Trim(SELDTMTAB_RS("HIV"))) > 0 Then
                    CrystalReport1.Formulas(8) = "HIV = '" & DECRPT(Trim(SELDTMTAB_RS("HIV")), a) & "'"
                Else
                    CrystalReport1.Formulas(8) = "HIV = '" & "" & "'"
                End If
            End If
'***** STAFF
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Destination = crptToPrinter
            'CrystalReport1.Destination = crptToWindow
            If Trim(SELDTMTAB_RS("ABO")) <> "" Or Trim(SELDTMTAB_RS("RH")) <> "" Or Trim(SELDTMTAB_RS("ABS")) <> "" Or Trim(SELDTMTAB_RS("DAT")) <> "" Or Trim(SELDTMTAB_RS("IAT")) <> "" Or Trim(SELDTMTAB_RS("HBSAG")) <> "" Or Trim(SELDTMTAB_RS("HCV")) <> "" Or Trim(SELDTMTAB_RS("HIV")) <> "" Then
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
          'JUSTIN 02-09-04
         'End If
            'CrystalReport1.Action = 1
        Next i
    End If
    If cmb_option.Text = "HCV(ELISA)" Then
        SELDTMNO_QR.QueryTimeout = 150
        SELDTMTAB_QR.QueryTimeout = 300
        SELREQUIQR.QueryTimeout = 300
        SELDTMNO_QR.QueryTimeout = 300
        For i = 1 To MSFhcv.Rows - 1
            If Trim(MSFhcv.TextMatrix(i, 1)) <> "" Then
                pcount = pcount + 1
            End If
        Next i
        For i = 1 To pcount 'MSFhcv.Rows - 1
            If SELDTMTAB_FLAG = False Then
                SELDTMTAB_QR(0) = req_chglen(Trim(MSFrcs.TextMatrix(i, 1)))
                Set SELDTMTAB_RS = SELDTMTAB_QR.OpenResultset(1, 3)
                SELDTMTAB_FLAG = True
            Else
                SELDTMTAB_QR(0) = req_chglen(Trim(MSFrcs.TextMatrix(i, 1)))
                SELDTMTAB_RS.MoreResults
                SELDTMTAB_RS.Requery
            End If
            
       'JUSTIN  02-09-04 CHECKING FOR PRINTED RECORD
     
'       If UCase(Trim(SELDTMTAB_RS("HCV"))) <> "Y" Or IsNull(Trim(SELDTMTAB_RS("HCV"))) Then
'
'            PUPDDTMTAB_QR(0) = "Y"
'            PUPDDTMTAB_QR(1) = req_chglen(Trim(MSFrcs.TextMatrix(i, 1)))
'            PUPDDTMTAB_QR.Execute
            
            
            
            MOD_CODE = Mid(MSFhcv.TextMatrix(i, 1), 5, 2)
            VAR_REQNNO = req_chglen(MSFhcv.TextMatrix(i, 1))
            If VARREQFLAG = False Then
                VARREQQR(0) = MSFhcv.TextMatrix(i, 2)
                VARREQQR(1) = VAR_REQNNO
                Set VARREQRS = VARREQQR.OpenResultset(1, 3)
                VARREQFLAG = True
            Else
                VARREQQR(0) = MSFhcv.TextMatrix(i, 2)
                VARREQQR(1) = VAR_REQNNO
                VARREQRS.MoreResults
                VARREQRS.Requery
            End If
            If SELREQUIFLAG = False Then
                SELREQUIQR(0) = MSFhcv.TextMatrix(i, 2)
                SELREQUIQR(1) = MOD_CODE
                'SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELREQUIQR(2) = "T"
                SELREQUIQR(3) = "F"
                Set SELREQUIRS = SELREQUIQR.OpenResultset(1, 3)
                SELREQUIFLAG = True
            Else
                SELREQUIQR(0) = MSFhcv.TextMatrix(i, 2)
                SELREQUIQR(1) = MOD_CODE
                'SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELREQUIQR(2) = "T"
                SELREQUIQR(3) = "F"
                SELREQUIRS.MoreResults
                SELREQUIRS.Requery
            End If
            If SELREQUIRS.RowCount > 0 Then
                prepath = ""
                While Not SELREQUIRS.EOF
                    If Trim(SELREQUIRS("reqno")) <> Trim(VAR_REQNNO) Then
                        prepath = prepath & req_chglen(SELREQUIRS("REQNO")) & " "
                    End If
                    SELREQUIRS.MoveNext
                Wend
            End If
            'Dim pcount As Integer
            pcount = 0
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = MSFhcv.TextMatrix(i, 2)
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = MSFhcv.TextMatrix(i, 2)
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If Opt_patient.Value = True Then
                str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtm.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
                End If
                CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
                If SELDTMNO_RS.EOF Then
                    CrystalReport1.Formulas(3) = "dtmno = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(3) = "dtmno = '" & SELDTMNO_RS("dtmno") & "'"
                End If
                a = CInt(SELDTMTAB_RS("hbsagkey"))
                If Len(Trim(SELDTMTAB_RS("HBSAG"))) > 0 Then
                    CrystalReport1.Formulas(4) = "HBSAG = '" & DECRPT(Trim(SELDTMTAB_RS("hbsag")), a) & "'"
                Else
                    CrystalReport1.Formulas(4) = "HBSAG = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HCVKEY"))
                If Len(Trim(SELDTMTAB_RS("HCV"))) > 0 Then
                    CrystalReport1.Formulas(5) = "HCV = '" & DECRPT(Trim(SELDTMTAB_RS("HCV")), a) & "'"
                Else
                    CrystalReport1.Formulas(5) = "HCV = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HIVKEY"))
                If Len(Trim(SELDTMTAB_RS("HIV"))) > 0 Then
                    CrystalReport1.Formulas(6) = "HIV = '" & DECRPT(Trim(SELDTMTAB_RS("HIV")), a) & "'"
                Else
                    CrystalReport1.Formulas(6) = "HIV = '" & "" & "'"
                End If
            End If
'**** STAFF
            If opt_staff.Value = True Then
                If SELEMPBASINFOFLAG = False Then
                    SELEMPBASINFO_QR(0) = MSFhcv.TextMatrix(i, 2)
                    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
                    SELEMPBASINFOFLAG = True
                Else
                    SELEMPBASINFO_QR(0) = MSFhcv.TextMatrix(i, 2)
                    SELEMPBASINFO_RS.MoreResults
                    SELEMPBASINFO_RS.Requery
                End If
                If REQ2_QRFLAG = False Then
                    req2_qr(0) = req_chglen(MSFhcv.TextMatrix(i, 1))
                    Set req2_rs = req2_qr.OpenResultset(1, 3)
                    REQ2_QRFLAG = True
                Else
                    req2_qr(0) = req_chglen(MSFhcv.TextMatrix(i, 1))
                    req2_rs.MoreResults
                    req2_rs.Requery
                End If
                str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(MSFhcv.TextMatrix(i, 1)) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtmstaff.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & MSFhcv.TextMatrix(i, 1) & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
                End If
                CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(3) = "NAM= '" & Trim(txtcasename.Text) & "'"
                CrystalReport1.Formulas(4) = "DOB= '" & DateDiff("YYYY", SELEMPBASINFO_RS("dob"), req2_rs("FREPORTDT")) & "'"
                CrystalReport1.Formulas(5) = "STDOB= '" & SELEMPBASINFO_RS("gender") & "'"
                a = CInt(SELDTMTAB_RS("hbsagkey"))
                If Len(Trim(SELDTMTAB_RS("HBSAG"))) > 0 Then
                    CrystalReport1.Formulas(6) = "HBSAG = '" & DECRPT(Trim(SELDTMTAB_RS("hbsag")), a) & "'"
                Else
                    CrystalReport1.Formulas(6) = "HBSAG = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HCVKEY"))
                If Len(Trim(SELDTMTAB_RS("HCV"))) > 0 Then
                    CrystalReport1.Formulas(7) = "HCV = '" & DECRPT(Trim(SELDTMTAB_RS("HCV")), a) & "'"
                Else
                    CrystalReport1.Formulas(7) = "HCV = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HIVKEY"))
                If Len(Trim(SELDTMTAB_RS("HIV"))) > 0 Then
                    CrystalReport1.Formulas(8) = "HIV = '" & DECRPT(Trim(SELDTMTAB_RS("HIV")), a) & "'"
                Else
                    CrystalReport1.Formulas(8) = "HIV = '" & "" & "'"
                End If
            End If
'***** STAFF
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Destination = crptToPrinter
            'CrystalReport1.Destination = crptToWindow
            If Trim(SELDTMTAB_RS("ABO")) <> "" Or Trim(SELDTMTAB_RS("RH")) <> "" Or Trim(SELDTMTAB_RS("ABS")) <> "" Or Trim(SELDTMTAB_RS("DAT")) <> "" Or Trim(SELDTMTAB_RS("IAT")) <> "" Or Trim(SELDTMTAB_RS("HBSAG")) <> "" Or Trim(SELDTMTAB_RS("HCV")) <> "" Or Trim(SELDTMTAB_RS("HIV")) <> "" Then
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
          'JUSTIN 02-09-04
          'End If
        Next i
    End If
    If cmb_option.Text = "HIV(ELISA)" Then
        SELDTMNO_QR.QueryTimeout = 150
        SELDTMTAB_QR.QueryTimeout = 300
        SELREQUIQR.QueryTimeout = 300
        SELDTMNO_QR.QueryTimeout = 300
        For i = 1 To MSFhiv.Rows - 1
            If Trim(MSFhiv.TextMatrix(i, 1)) <> "" Then
                pcount = pcount + 1
            End If
        Next i
        For i = 1 To pcount 'MSFhiv.Rows - 1
            If SELDTMTAB_FLAG = False Then
                SELDTMTAB_QR(0) = req_chglen(Trim(MSFrcs.TextMatrix(i, 1)))
                Set SELDTMTAB_RS = SELDTMTAB_QR.OpenResultset(1, 3)
                SELDTMTAB_FLAG = True
            Else
                SELDTMTAB_QR(0) = req_chglen(Trim(MSFrcs.TextMatrix(i, 1)))
                SELDTMTAB_RS.MoreResults
                SELDTMTAB_RS.Requery
            End If
            
       'JUSTIN  02-09-04 CHECKING FOR PRINTED RECORD
     
'       If UCase(Trim(SELDTMTAB_RS("HCV"))) <> "Y" Or IsNull(Trim(SELDTMTAB_RS("HCV"))) Then
'
'            PUPDDTMTAB_QR(0) = "Y"
'            PUPDDTMTAB_QR(1) = req_chglen(Trim(MSFrcs.TextMatrix(i, 1)))
'            PUPDDTMTAB_QR.Execute
            
            MOD_CODE = Mid(MSFhiv.TextMatrix(i, 1), 5, 2)
            VAR_REQNNO = req_chglen(MSFhiv.TextMatrix(i, 1))
            If VARREQFLAG = False Then
                VARREQQR(0) = MSFhiv.TextMatrix(i, 2)
                VARREQQR(1) = VAR_REQNNO
                Set VARREQRS = VARREQQR.OpenResultset(1, 3)
                VARREQFLAG = True
            Else
                VARREQQR(0) = MSFhiv.TextMatrix(i, 2)
                VARREQQR(1) = VAR_REQNNO
                VARREQRS.MoreResults
                VARREQRS.Requery
            End If
            If SELREQUIFLAG = False Then
                SELREQUIQR(0) = MSFhiv.TextMatrix(i, 2)
                SELREQUIQR(1) = MOD_CODE
                'SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELREQUIQR(2) = "T"
                SELREQUIQR(3) = "F"
                Set SELREQUIRS = SELREQUIQR.OpenResultset(1, 3)
                SELREQUIFLAG = True
            Else
                SELREQUIQR(0) = MSFhiv.TextMatrix(i, 2)
                SELREQUIQR(1) = MOD_CODE
                'SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
                SELREQUIQR(2) = "T"
                SELREQUIQR(3) = "F"
                SELREQUIRS.MoreResults
                SELREQUIRS.Requery
            End If
            If SELREQUIRS.RowCount > 0 Then
                prepath = ""
                While Not SELREQUIRS.EOF
                    If Trim(SELREQUIRS("reqno")) <> Trim(VAR_REQNNO) Then
                        prepath = prepath & req_chglen(SELREQUIRS("REQNO")) & " "
                    End If
                    SELREQUIRS.MoveNext
                Wend
            End If
            'Dim pcount As Integer
            pcount = 0
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = MSFhiv.TextMatrix(i, 2)
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = MSFhiv.TextMatrix(i, 2)
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If Opt_patient.Value = True Then
                str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtm.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
                End If
                CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
                If SELDTMNO_RS.EOF Then
                    CrystalReport1.Formulas(3) = "dtmno = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(3) = "dtmno = '" & SELDTMNO_RS("dtmno") & "'"
                End If
                a = CInt(SELDTMTAB_RS("hbsagkey"))
                If Len(Trim(SELDTMTAB_RS("HBSAG"))) > 0 Then
                    CrystalReport1.Formulas(4) = "HBSAG = '" & DECRPT(Trim(SELDTMTAB_RS("hbsag")), a) & "'"
                Else
                    CrystalReport1.Formulas(4) = "HBSAG = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HCVKEY"))
                If Len(Trim(SELDTMTAB_RS("HCV"))) > 0 Then
                    CrystalReport1.Formulas(5) = "HCV = '" & DECRPT(Trim(SELDTMTAB_RS("HCV")), a) & "'"
                Else
                    CrystalReport1.Formulas(5) = "HCV = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HIVKEY"))
                If Len(Trim(SELDTMTAB_RS("HIV"))) > 0 Then
                    CrystalReport1.Formulas(6) = "HIV = '" & DECRPT(Trim(SELDTMTAB_RS("HIV")), a) & "'"
                Else
                    CrystalReport1.Formulas(6) = "HIV = '" & "" & "'"
                End If
            End If
'**** STAFF
            If opt_staff.Value = True Then
                If SELEMPBASINFOFLAG = False Then
                    SELEMPBASINFO_QR(0) = MSFhiv.TextMatrix(i, 2)
                    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
                    SELEMPBASINFOFLAG = True
                Else
                    SELEMPBASINFO_QR(0) = MSFhiv.TextMatrix(i, 2)
                    SELEMPBASINFO_RS.MoreResults
                    SELEMPBASINFO_RS.Requery
                End If
                If REQ2_QRFLAG = False Then
                    req2_qr(0) = req_chglen(MSFhiv.TextMatrix(i, 1))
                    Set req2_rs = req2_qr.OpenResultset(1, 3)
                    REQ2_QRFLAG = True
                Else
                    req2_qr(0) = req_chglen(MSFhiv.TextMatrix(i, 1))
                    req2_rs.MoreResults
                    req2_rs.Requery
                End If
                str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(MSFhiv.TextMatrix(i, 1)) & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtmstaff.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & MSFhiv.TextMatrix(i, 1) & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
                End If
                CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(3) = "NAM= '" & Trim(txtcasename.Text) & "'"
                CrystalReport1.Formulas(4) = "DOB= '" & DateDiff("YYYY", SELEMPBASINFO_RS("dob"), req2_rs("FREPORTDT")) & "'"
                CrystalReport1.Formulas(5) = "STDOB= '" & SELEMPBASINFO_RS("gender") & "'"
                a = CInt(SELDTMTAB_RS("hbsagkey"))
                If Len(Trim(SELDTMTAB_RS("HBSAG"))) > 0 Then
                    CrystalReport1.Formulas(6) = "HBSAG = '" & DECRPT(Trim(SELDTMTAB_RS("hbsag")), a) & "'"
                Else
                    CrystalReport1.Formulas(6) = "HBSAG = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HCVKEY"))
                If Len(Trim(SELDTMTAB_RS("HCV"))) > 0 Then
                    CrystalReport1.Formulas(7) = "HCV = '" & DECRPT(Trim(SELDTMTAB_RS("HCV")), a) & "'"
                Else
                    CrystalReport1.Formulas(7) = "HCV = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HIVKEY"))
                If Len(Trim(SELDTMTAB_RS("HIV"))) > 0 Then
                    CrystalReport1.Formulas(8) = "HIV = '" & DECRPT(Trim(SELDTMTAB_RS("HIV")), a) & "'"
                Else
                    CrystalReport1.Formulas(8) = "HIV = '" & "" & "'"
                End If
            End If
'***** STAFF
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Destination = crptToPrinter
            'CrystalReport1.Destination = crptToWindow
            
            If Trim(SELDTMTAB_RS("ABO")) <> "" Or Trim(SELDTMTAB_RS("RH")) <> "" Or Trim(SELDTMTAB_RS("ABS")) <> "" Or Trim(SELDTMTAB_RS("DAT")) <> "" Or Trim(SELDTMTAB_RS("IAT")) <> "" Or Trim(SELDTMTAB_RS("HBSAG")) <> "" Or Trim(SELDTMTAB_RS("HCV")) <> "" Or Trim(SELDTMTAB_RS("HIV")) <> "" Then
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
          'JUSTIN 02-09-04
          'End If
        Next i
    End If
End If
If OPT_CASENO.Value = True Then
    If Trim(Com_REQNO.Text) = "" Then
        MsgBox "Requisition No. Field is Empty !!", vbInformation, "DTM Lab Report"
        Exit Sub
    End If
    VAR_REQNNO = req_chglen(Com_REQNO.Text)
    If SELDTMTAB_FLAG = False Then
        SELDTMTAB_QR(0) = VAR_REQNNO
        Set SELDTMTAB_RS = SELDTMTAB_QR.OpenResultset(1, 3)
        SELDTMTAB_FLAG = True
    Else
        SELDTMTAB_QR(0) = VAR_REQNNO
        SELDTMTAB_RS.MoreResults
        SELDTMTAB_RS.Requery
    End If
    
    'JUSTIN  02-09-04 CHECKING FOR PRINTED RECORD
     
'    If UCase(Trim(SELDTMTAB_RS("HCV"))) <> "Y" Or IsNull(Trim(SELDTMTAB_RS("HCV"))) Then
'
'        PUPDDTMTAB_QR(0) = "Y"
'        PUPDDTMTAB_QR(1) = VAR_REQNNO
'        PUPDDTMTAB_QR.Execute
    
    
    i = InStr(Trim(com_modality.Text), "$")
    MOD_CODE = Mid(Trim(com_modality.Text), i + 1, 2)
    
    If VARREQFLAG = False Then
        VARREQQR(0) = msk_caseno.FormattedText
        VARREQQR(1) = req_chglen(Com_REQNO.Text)
        Set VARREQRS = VARREQQR.OpenResultset(1, 3)
        VARREQFLAG = True
    Else
        VARREQQR(0) = msk_caseno.FormattedText
        VARREQQR(1) = VAR_REQNNO
        VARREQRS.MoreResults
        VARREQRS.Requery
    End If
    If SELREQUIFLAG = False Then
        SELREQUIQR(0) = msk_caseno.FormattedText
        SELREQUIQR(1) = MOD_CODE
        'SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
        SELREQUIQR(2) = "T"
        SELREQUIQR(3) = "F"
        Set SELREQUIRS = SELREQUIQR.OpenResultset(1, 3)
        SELREQUIFLAG = True
    Else
        SELREQUIQR(0) = msk_caseno.FormattedText
        SELREQUIQR(1) = MOD_CODE
        'SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
        SELREQUIQR(2) = "T"
        SELREQUIQR(3) = "F"
        SELREQUIRS.MoreResults
        SELREQUIRS.Requery
    End If
    If SELREQUIRS.RowCount > 0 Then
        prepath = ""
        While Not SELREQUIRS.EOF
            If Trim(SELREQUIRS("reqno")) <> VAR_REQNNO Then
                prepath = prepath & req_chglen(SELREQUIRS("REQNO")) & " "
            End If
            SELREQUIRS.MoveNext
        Wend
     End If
    pcount = 0
    If SELDTMNO_FLAG = False Then
        SELDTMNO_QR(0) = msk_caseno.FormattedText
        Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
        SELDTMNO_FLAG = True
    Else
        SELDTMNO_QR(0) = msk_caseno.FormattedText
        SELDTMNO_RS.MoreResults
        SELDTMNO_RS.Requery
    End If
    If Opt_patient.Value = True Then
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtm.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & VAR_REQNNO & "'"
        If Trim(prepath) = "" Then
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
        End If
        CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
        If SELDTMNO_RS.EOF Then
            CrystalReport1.Formulas(3) = "dtmno = '" & "" & "'"
        Else
            CrystalReport1.Formulas(3) = "dtmno = '" & SELDTMNO_RS("dtmno") & "'"
        End If
        a = CInt(SELDTMTAB_RS("hbsagkey"))
        If Len(Trim(SELDTMTAB_RS("HBSAG"))) > 0 Then
            CrystalReport1.Formulas(4) = "HBSAG = '" & DECRPT(Trim(SELDTMTAB_RS("hbsag")), a) & "'"
        Else
            CrystalReport1.Formulas(4) = "HBSAG = '" & "" & "'"
        End If
        a = CInt(SELDTMTAB_RS("HCVKEY"))
        If Len(Trim(SELDTMTAB_RS("HCV"))) > 0 Then
            CrystalReport1.Formulas(5) = "HCV = '" & DECRPT(Trim(SELDTMTAB_RS("HCV")), a) & "'"
        Else
            CrystalReport1.Formulas(5) = "HCV = '" & "" & "'"
        End If
        a = CInt(SELDTMTAB_RS("HIVKEY"))
        If Len(Trim(SELDTMTAB_RS("HIV"))) > 0 Then
            CrystalReport1.Formulas(6) = "HIV = '" & DECRPT(Trim(SELDTMTAB_RS("HIV")), a) & "'"
        Else
            CrystalReport1.Formulas(6) = "HIV = '" & "" & "'"
        End If
    End If
'**** STAFF
    If opt_staff.Value = True Then
        If SELEMPBASINFOFLAG = False Then
            SELEMPBASINFO_QR(0) = msk_caseno.Text
            Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
            SELEMPBASINFOFLAG = True
        Else
            SELEMPBASINFO_QR(0) = msk_caseno.Text
            SELEMPBASINFO_RS.MoreResults
            SELEMPBASINFO_RS.Requery
        End If
        If REQ2_QRFLAG = False Then
            req2_qr(0) = req_chglen(Trim(Com_REQNO.Text))
            Set req2_rs = req2_qr.OpenResultset(1, 3)
            REQ2_QRFLAG = True
        Else
            req2_qr(0) = req_chglen(Trim(Com_REQNO.Text))
            req2_rs.MoreResults
            req2_rs.Requery
        End If
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(Trim(Com_REQNO.Text)) & "'"
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtmstaff.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(Com_REQNO.Text) & "'"
        If Trim(prepath) = "" Then
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
        End If
        CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(3) = "NAM= '" & Trim(txtcasename.Text) & "'"
        CrystalReport1.Formulas(4) = "DOB= '" & DateDiff("YYYY", SELEMPBASINFO_RS("dob"), req2_rs("FREPORTDT")) & "'"
        CrystalReport1.Formulas(5) = "STDOB= '" & SELEMPBASINFO_RS("gender") & "'"
        a = CInt(SELDTMTAB_RS("hbsagkey"))
        If Len(Trim(SELDTMTAB_RS("HBSAG"))) > 0 Then
            CrystalReport1.Formulas(6) = "HBSAG = '" & DECRPT(Trim(SELDTMTAB_RS("hbsag")), a) & "'"
        Else
            CrystalReport1.Formulas(6) = "HBSAG = '" & "" & "'"
        End If
        a = CInt(SELDTMTAB_RS("HCVKEY"))
        If Len(Trim(SELDTMTAB_RS("HCV"))) > 0 Then
            CrystalReport1.Formulas(7) = "HCV = '" & DECRPT(Trim(SELDTMTAB_RS("HCV")), a) & "'"
        Else
            CrystalReport1.Formulas(7) = "HCV = '" & "" & "'"
        End If
        a = CInt(SELDTMTAB_RS("HIVKEY"))
        If Len(Trim(SELDTMTAB_RS("HIV"))) > 0 Then
            CrystalReport1.Formulas(8) = "HIV = '" & DECRPT(Trim(SELDTMTAB_RS("HIV")), a) & "'"
        Else
            CrystalReport1.Formulas(8) = "HIV = '" & "" & "'"
        End If
    End If
    '***** STAFF
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Destination = crptToWindow
'    CRYSTALREPORT1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
   'JUSTIN 02-09-04
   'End If
End If
Exit Sub
prtmsg: MsgBox Error(Err)
'txt_msg.Text = ""
btnenabl ("print")
Exit Sub
MSG:
MsgBox Error(Err)

End Sub

Private Sub com_dept_GotFocus()
'On Error GoTo MSG:
'Dim ind_cbodept As Integer
'If Trim(cmb_dept) = "" Then
'    For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
'        If Trim(cmb_dept.List(ind_cbodept)) = "HAEMATO ONCOLOGY LABORATORY$UZZ" Then
'            cmb_dept.ListIndex = ind_cbodept
'        End If
'    Next ind_cbodept
'    End If
'Exit Sub
'MSG:
'MsgBox Error(Err)


If sel_departmentflag = False Then
       ' selma106_qr(0) = "NZZ"
    Set sel_departmentrs = SELMA106_QR.OpenResultset(1, 3)
    sel_departmentflag = True
Else
    'selma106_qr(0) = "NZZ"
    sel_departmentrs.MoreResults
    sel_departmentrs.Requery
End If

'Commented by Shraddha on 31/08/09
'If Not sel_departmentrs.EOF Then
'    com_dept.clear
'    com_dept.AddItem sel_departmentrs("dept_name")
'    com_dept.Text = Trim(sel_departmentrs("dept_name"))
'End If
'Till here

'        If SEL_MODALITYFLAG1 = False Then
'            'SELRMODALITY_QR(0) = "NZZ"
'            Set sel_modalityrs = SELRMODALITY_QR.OpenResultset(1, 3)
'            SEL_MODALITYFLAG1 = True
'        Else
'            'SELRMODALITY_QR(0) = "NZZ"
'            sel_modalityrs.MoreResults
'            sel_modalityrs.Requery
'        End If
'        If Not sel_modalityrs.EOF Then
'            com_modality.clear
'            Do While Not sel_modalityrs.EOF
'                com_modality.AddItem Trim(sel_modalityrs(1)) & "$" & Trim(sel_modalityrs(0))
'                sel_modalityrs.MoveNext
'            Loop
'        End If
End Sub

Private Sub com_modality_GotFocus()
If SEL_MODALITYFLAG1 = False Then
    'SELRMODALITY_QR(0) = "NZZ"
    Set sel_modalityrs = SELRMODALITY_QR.OpenResultset(1, 3)
    SEL_MODALITYFLAG1 = True
Else
    'SELRMODALITY_QR(0) = "NZZ"
    sel_modalityrs.MoreResults
    sel_modalityrs.Requery
End If

'Commented by Shraddha on 31/08/2009
'If Not sel_modalityrs.EOF Then
'    com_modality.clear
'    Do While Not sel_modalityrs.EOF
'        com_modality.AddItem Trim(sel_modalityrs(1)) & "$" & Trim(sel_modalityrs(0))
'        com_modality.Text = Trim(sel_modalityrs(1)) & "$" & Trim(sel_modalityrs(0))
'        sel_modalityrs.MoveNext
'    Loop
'End If
'Till here Shraddha
End Sub

Private Sub com_reqno_Click()
If SELDMLPROCE_FLAG = False Then
    SELDMLPROCE_QR(0) = req_chglen(Trim(Com_REQNO.Text))
    Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
    SELDMLPROCE_FLAG = True
Else
    SELDMLPROCE_QR(0) = req_chglen(Trim(Com_REQNO.Text))
    SELDMLPROCE_RS.MoreResults
    SELDMLPROCE_RS.Requery
End If
If Not SELDMLPROCE_RS.EOF Then
    cmb_option.clear
    Do While Not SELDMLPROCE_RS.EOF
        If SELDMLPROCE_RS("PRO_CODE") = "H001" Or SELDMLPROCE_RS("PRO_CODE") = "H007" Or SELDMLPROCE_RS("PRO_CODE") = "H006" Then
            cmb_option.AddItem "RCS"            '
'        ElseIf SELDMLPROCE_RS("PRO_CODE") = "H012" Then
'            cmb_option.AddItem "HBsAg(ELISA)"
'        ElseIf SELDMLPROCE_RS("PRO_CODE") = "H014" Then
'            cmb_option.AddItem "HCV(ELISA)"
'        ElseIf SELDMLPROCE_RS("PRO_CODE") = "H013" Then
'            cmb_option.AddItem "HIV(ELISA)"
        End If
        SELDMLPROCE_RS.MoveNext
    Loop
End If
' ADD CODE BY SANDEEP
If SELEDREQSTAFFD_FLAG = False Then
        SELEDREQSTAFFD_QR(0) = msk_caseno.Text
        SELEDREQSTAFFD_QR(1) = req_chglen(Trim(Com_REQNO.Text))
        Set SELEDREQSTAFFD_RS = SELEDREQSTAFFD_QR.OpenResultset(1, 3)
        SELEDREQSTAFFD_FLAG = True
Else
        SELEDREQSTAFFD_RS.MoreResults
        SELEDREQSTAFFD_RS.Requery
End If

'****************** Added on 08/06/09 by Shraddha **************
SELEDREQSTAFFD_QR(0) = msk_caseno.Text
SELEDREQSTAFFD_QR(1) = req_chglen(Trim(Com_REQNO.Text))
Set SELEDREQSTAFFD_RS = SELEDREQSTAFFD_QR.OpenResultset(1, 3)
'***************************************************************
        If SELEDREQSTAFFD_RS.RowCount > 0 Then
            txtcasename.Text = SELEDREQSTAFFD_RS("PNAME")
        End If
'TXT_COMM.Visible = True

cmb_option.SetFocus

End Sub

Private Sub cmd_abs_Click()
If MODIFYFLAG = True Then
    MSFrcs.row = 1
    MSFrcs.col = 5
    cmb_abs.ListIndex = -1
    cmb_abs.Top = MSFrcs.CellTop + MSFrcs.Top
    cmb_abs.Visible = True
    cmb_abo.Visible = False
    cmb_rh.Visible = False
    cmb_dat.Visible = False
    cmb_iat.Visible = False
    cmb_abs.SetFocus
End If
End Sub

Private Sub cmd_anti_hcv_Click()
If MODIFYFLAG = True Then
    MSFhcv.row = 1
    MSFhcv.col = 3
    cmb_anti_hcv.ListIndex = -1
    cmb_anti_hcv.Top = MSFhcv.CellTop + MSFhcv.Top
    cmb_anti_hcv.Visible = True
    cmb_anti_hcv.SetFocus
End If
End Sub

Private Sub cmd_anti_hiv_Click()
If MODIFYFLAG = True Then
    MSFhiv.row = 1
    MSFhiv.col = 3
    cmb_anti_hiv.ListIndex = -1
    cmb_anti_hiv.Top = MSFhiv.CellTop + MSFhiv.Top
    cmb_anti_hiv.Visible = True
    cmb_anti_hiv.SetFocus
End If
End Sub
Private Sub cmd_back_Click()
CMDNEW4VIEW = False
Unload Me
FLDENABL ("BACK")
DIS_MAIN.show
End Sub
Private Sub cmd_bldgrp_Click()
If MODIFYFLAG = True Then
    MSFrcs.row = 1
    MSFrcs.col = 3
    cmb_abo.ListIndex = -1
    cmb_abo.Top = MSFrcs.CellTop + MSFrcs.Top
    cmb_abo.Visible = True
    MSFrcs.col = 4
    cmb_rh.ListIndex = -1
    cmb_rh.Top = MSFrcs.CellTop + MSFrcs.Top
    cmb_rh.Visible = True
    cmb_dat.Visible = False
    cmb_iat.Visible = False
    cmb_abs.Visible = False
    cmb_abo.SetFocus
End If
End Sub
Private Sub cmd_cancel_Click()
CMDNEW4VIEW = False
If cmb_option.Text = Trim("RCS") Then
    MSFrcs.clear
    MSFrcs.FormatString = "SN|^   Requisition No.     |^  Case No. |^      ABO       |^      Rh     |^    AbS      |^    DAT     |^    IAT       |^   DTM No.  "
ElseIf cmb_option.Text = Trim("HBsAg(ELISA)") Then
    MSFhbsAg.clear
    MSFhbsAg.FormatString = "S.No |^           Requisition No.        |^         Case No.       |^            HBsAg         |^      DTM No.   "
ElseIf cmb_option.Text = Trim("HCV(ELISA)") Then
    MSFhcv.clear
    MSFhcv.FormatString = "S.No |^            Requisition No.          |^  Case No.   |^      Anti - HCV (ELISA)   |^    DTM No.     "
ElseIf cmb_option.Text = Trim("HIV(ELISA)") Then
    MSFhiv.clear
    MSFhiv.FormatString = "S.No |^           Requisition No.       |^  Case No.   |^       Anti- HIV (ELISA)     |^       DTM No.     "
End If
btnenabl ("INI")
FLDENABL ("CANCEL")
field_clear
BCCLEAR
TXT_CLEAR
viewflag = False
NEWFLAG = False
diffcountflag = False
usr = ""
SSTab1.TabVisible(0) = True
SSTab1.TabVisible(1) = False
SSTab1.TabVisible(2) = False
SSTab1.TabVisible(3) = False
SSTab1.Enabled = False
msk_caseno.Enabled = True
Com_REQNO.Enabled = True
OPT_ALL.Value = False
OPT_CASENO.Value = False
msk_fromdt.Enabled = True
msk_todt.Enabled = True
End Sub
Private Sub cmd_dat_Click()
If MODIFYFLAG = True Then
    MSFrcs.row = 1
    MSFrcs.col = 6
    cmb_dat.ListIndex = -1
    cmb_dat.Top = MSFrcs.CellTop + MSFrcs.Top
    cmb_abo.Visible = False
    cmb_rh.Visible = False
    cmb_dat.Visible = True
    cmb_iat.Visible = False
    cmb_abs.Visible = False
    cmb_dat.SetFocus
End If
End Sub
Private Sub cmd_hbsag_Click()
If MODIFYFLAG = True Then
    MSFhbsAg.row = 1
    MSFhbsAg.col = 3
    cmb_hbsag.ListIndex = -1
    cmb_hbsag.Top = MSFhbsAg.CellTop + MSFhbsAg.Top
    cmb_hbsag.Visible = True
    cmb_hbsag.SetFocus
End If
End Sub
Private Sub cmd_iat_Click()
If MODIFYFLAG = True Then
    MSFrcs.row = 1
    MSFrcs.col = 7
    cmb_iat.ListIndex = -1
    cmb_iat.Top = MSFrcs.CellTop + MSFrcs.Top
    cmb_iat.Visible = False
    cmb_rh.Visible = False
    cmb_dat.Visible = False
    cmb_abs.Visible = False
    cmb_iat.Visible = True
    cmb_iat.SetFocus
End If
End Sub

Private Sub cmd_modify_Click()
CMDNEW4VIEW = False
On Error GoTo MSG:
flagsave = 2
btnenabl ("NEW")
FLDENABL ("MODIFY")
RPTTIME_SEL ("MODIFY")
MODIFYFLAG = True
flagsave = 2
Exit Sub
MSG:
MsgBox Err.Description, vbCritical, "DTM"
Exit Sub
End Sub

Private Sub cmd_new_Click()
CMDNEW4VIEW = True
Dim L1 As Integer
'On Error GoTo MSG:
btnenabl ("NEW")
RPTTIME_SEL ("NEW")
RPTTIME_SEL ("NEW")
field_clear
TXT_CLEAR
If DATETIMEFLAG = False Then
   datetime
   DATETIMEFLAG = True
Else
   datetime
End If
txt_date.Text = DATE11
txt_time.Text = time1
cmb_enterby.Text = cmb_enterby.List(0)
'***** Shraddha Pandit 10/02/10 *****
com_dept.Text = "TRANSFUSION MEDICINE"
com_modality.Text = "SCREENING SERVICES$SS"
'************************************
cmb_enterby.Enabled = False
For L1 = 0 To cmb_finalby.ListCount
    If Trim(cmb_enterby.List(0)) = cmb_finalby.List(L1) Then
        cmb_finalby.Text = cmb_finalby.List(L1)
    End If
Next L1
'cmb_finalby.Text = cmb_finalby.List(0)
BCCLEAR
Opt_patient.Value = True
flagsave = 1
viewflag = False
MODIFYFLAG = True
NEWFLAG = True
FLDENABL ("ENA")
'MaskEdBox1.Enabled = True
    'SUNDEEP dR.NARAYAN
    SSTab1.TabVisible(0) = True
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    'SUNDEEP dR.NARAYAN
'txtusername.SetFocus
'cmb_dept.SetFocus
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_save_Click()
CMDNEW4VIEW = False
If Len(Trim(cmb_enterby.Text)) = 0 Then
    MsgBox "The Enter by field is EMPTY.", vbInformation, "DTM Lab"
    cmb_enterby.SetFocus
Exit Sub
End If
btnenabl ("SAVE")
FLDENABL ("SAVE")
assign
'cmd_commit.Enabled = True
End Sub

Private Sub cmd_view_Click()
CMDNEW4VIEW = False

'On Error GoTo MSG:

viewflag = True



NEWFLAG = False
flagsave = 2
MODIFYFLAG = False
btnenabl ("VIEW")
RPTTIME_SEL ("VIEW")
FLDENABL ("MOD")
field_clear
TXT_CLEAR
BCCLEAR
If DATETIMEFLAG = False Then
   datetime
   DATETIMEFLAG = True
Else
    datetime
End If
txt_date.Text = DATE11
txt_time.Text = time1

'***** Shraddha Pandit 10/02/10 *****
com_dept.Text = "TRANSFUSION MEDICINE"
com_modality.Text = "SCREENING SERVICES$SS"
'************************************

Opt_patient.Value = True
'MaskEdBox1.Enabled = True
VIS = ""
'txtusername.SetFocus
com_dept.SetFocus
Exit Sub
MSG:
MsgBox Err.Description
End Sub
Sub RPTTIME_SEL(ByVal VIS As String)
If VIS = "VIEW" Or VIS = "NEW" Then
Label11.Visible = True
Label12.Visible = True
FTIME.Visible = True
TTIME.Visible = True
Else
Label11.Visible = False
Label12.Visible = False
FTIME.Visible = False
TTIME.Visible = False
End If
End Sub
Private Sub com_modality_LostFocus()
Dim i
i = InStr(Trim(com_modality.Text), "$")
MOD_CODE = Mid(Trim(com_modality.Text), i + 1, 2)
End Sub

Private Sub Com_RENO_Click()
If OPT_ALL.Value = True Then
    If REQ2_QRFLAG = False Then
        req2_qr(0) = Trim(Com_REQNO.Text)
        Set req2_rs = req2_qr.OpenResultset(1, 3)
        msk_caseno.Text = req2_rs(0)
        txtcasename.Text = req2_rs(9)
        REQ2_QRFLAG = True
    Else
        req2_qr(0) = Trim(Com_REQNO.Text)
        req2_rs.MoreResults
        req2_rs.Requery
        msk_caseno.Text = req2_rs(0)
        txtcasename.Text = req2_rs(9)
    End If
End If
End Sub

Private Sub Command1_Click()
Dim var_date
var_date = Format(DATE11, "DD/MM/YYYY")
var_date = Year(var_date) & "," & month(var_date) & "," & Day(var_date)
'{HEMDAILYLF.SAMRECDATE}= DATE(" & var_date & ")"
CrystalReport1.SelectionFormula = "{DTMREPORTF.PREPORTDT}=DATE(" & Trim(var_date) & ") AND {DTMREPORTF.PREPORTBY}='" & Trim(cmb_enterby.Text) & "'"
CrystalReport1.ReportFileName = "C:\DISREPORTSEClf\DTMDAILYREPORT.rpt"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=TMH;PWD=TMH;DSN=DISDSN"
CrystalReport1.Action = 1
' CrystalReport1.Destination = crptToPrinter
End Sub

Private Sub Form_Load()
'On Error GoTo MSG:

NEWFLAG = False
viewflag = False
FLDENABL ("DIS")
datetime
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1
If DTMFLAG = False Then
    DTMFLAG = True
    '*********SECURITY**********************************

str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME1DTM", str)
USRNAME_QR(0) = UID
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)
If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = CInt(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    LOGINCOMMIT = Trim(USRNAME_RS("COMMIT")) 'add for chk commit authority.
    cmb_enterby.AddItem LOGINUSRNAME
    cmb_enterby.Text = Trim(LOGINUSRNAME)
    cmb_enterby.Text = cmb_enterby.List(0)
Else
    MsgBox "The Security Detials of this User is not available!!", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If

'Set REPWTRIRS = rdb.OpenResultset("SELECT * FROM DISLIB.DISLNKUSR ORDER BY USR_ID", 1, 1)
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE DEPT_CODE='NZZ' ORDER BY USERNAME"
Set REPWTRN_QR = rdb.CreateQuery("REPWTRN_STRDDO", str)
'REPWTRN_QR(0) = Trim(USRDEPT)
Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
If REPWTRNRS.RowCount > 0 Then
    'cmb_enteryby.clear
    'cmb_addendumby.clear
    cmb_finalby.clear
    'cmb_provby.clear
    REPWTRNRS.MoveFirst
    While Not REPWTRNRS.EOF
        If ((REPWTRNRS("USRTYPE") <> 4) And (REPWTRNRS("USRTYPE") > 1)) Then
            'cmb_enteryby.AddItem Trim(REPWTRNRS("USERNAME"))
            cmb_finalby.AddItem Trim(REPWTRNRS("USERNAME"))
            'cmb_addendumby.AddItem Trim(REPWTRNRS("USERNAME"))
            'cmb_provby.AddItem Trim(REPWTRNRS("USERNAME"))
        End If
        REPWTRNRS.MoveNext
        cmb_finalby.Text = cmb_finalby.List(0)
    Wend
Else
    MsgBox "The Security Detials of this User is not available", vbInformation, "DIS"
    cmd_cancel_Click
    Exit Sub
End If
str = "select * from DISLIB.Dislnkusr where username = ?"
Set CURUSRDET_QR = rdb.CreateQuery("curusrdetq11DTM", str)
CURUSRDET_QR(0) = "ABC"
Set CURUSRDET_RS = CURUSRDET_QR.OpenResultset(1, 1)
'
'STR = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE=?"
'Set SELFINALBY_QR = rdb.CreateQuery("SELFINALBY_STR1DTM", STR)
'SELFINALBY_QR(0) = 6
'Set SELFINALBY_RS = SELFINALBY_QR.OpenResultset(1, 3)

'If Not SELFINALBY_RS.EOF Then
'    Do While Not SELFINALBY_RS.EOF
'       cmb_finalby.AddItem Trim(SELFINALBY_RS("USERNAME"))
'       SELFINALBY_RS.MoveNext
'    Loop
'End If
 str = "SELECT * FROM DISLIB.DMBCCOMM WHERE REQNO=?"
    Set SEL_DMBCCOMM_QR = rdb.CreateQuery("", str)
    
    str = "UPDATE DISLIB.DMBCCOMM SET COMMENTS=? WHERE REQNO=?"
    Set UPD_DMBCCOMM_qr = rdb.CreateQuery("", str)
        
 
  str = "INSERT INTO dislib.DMBCCOMM (REQNO,COMMENTS) VALUES (?,?)"
    Set INS_DMBCCOMM_QR = rdb.CreateQuery("", str)
  
      str = "SELECT * FROM DISLIB.DMBCCOMM WHERE REQNO=?"
    Set SEL_DMBCCOMM_QR = rdb.CreateQuery("", str)
    

'*********SECURITY**********************************
btnenabl ("INI")
    
    ''SELECT QUERY FOR DTMTAB TABLE
    str = "SELECT * FROM DISLIB.DTMTAB WHERE REQNO=?"
    Set SELDTMTAB_QR = rdb.CreateQuery("SELDTMTAB_STR", str)
    
    ''UPDATE QUERY FOR DTMTAB TABLE FOR PRINT CHECK JUSTIN 02-04-2004
     str = "UPDATE DISLIB.DTMTAB SET HCV=? WHERE REQNO=?"
     Set PUPDDTMTAB_QR = rdb.CreateQuery("PUPDDTMHIV_STR", str)
    
    ''UPDATE QUERY FOR DTMTAB TABLE
    str = "UPDATE DISLIB.DTMTAB SET HIV=? WHERE REQNO=?"
    Set UPDDTMTAB_QR = rdb.CreateQuery("UPDDTMHIV_STR", str)
    
    ''SELECT QUERY FOR DMULPROCE TABLE
    str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
    Set SELDMLPROCE_QR = rdb.CreateQuery("SELDMLPROCE_STR111", str)
    
    ''UPDATE QUERY FOR DTMNO TABLE
    str = "UPDATE DISLIB.DTMNO SET FLD1=? WHERE CASENO=?"
    Set UPDDTMNO_QR = rdb.CreateQuery("UPDDTMNO_STR111", str)
    
    ''SELECT QUERY FOR DTMNO TABLE
    str = "select * from dislib.dtmno where caseno=?"
    Set SELDTMNO_QR = rdb.CreateQuery("SELDTMNO_STR10111", str)
    
    ''UPDATE QUERY FOR DTMTAB TABLE
    str = "UPDATE DISLIB.DTMTAB SET ABO=?,RH=?,ABS=?,DAT=?,IAT=? WHERE CASENO=? AND REQNO=?"
    Set UPDDTMQRABO = rdb.CreateQuery("UPDDTMABOSTR", str)
    
    ''SELECT QUERY FOR DMBCCOMM TABLE
    str = "SELECT * FROM DISLIB.DMBCCOMM WHERE REQNO=?"
    Set SELDMBC_QR = rdb.CreateQuery("SELDMBC_STR", str)
  
    ''SELECT QUERY FOR MA106 TABLE
    str = "SELECT * FROM PATOOLSLIB.MA106 WHERE DEPT_CODE='NZZ' "
    Set SELMA106_QR = rdb.CreateQuery("SELMA106_dtm11", str)
    
    ''SELECT QUERY FOR RMODALITY TABLE
    str = "select * from rislib.rmodality where dept_code='NZZ' and modalityCD='SS'"
    Set SELRMODALITY_QR = rdb.CreateQuery("selmodality_dtm11", str)
    
    ''SELECT QUERY FOR RREQUISITN TABLE
    str = "select * from Rislib.RREQUISITN where caseno=? AND MODALITYCD=? AND REQNSTAT='E' ORDER BY FREPORTDT desc "
    Set req_qr = rdb.CreateQuery("reqDTM", str)
    
    ''SELECT QUERY FOR RREQUISITN TABLE
    str = "select * from rislib.rrequisitn where FREPORTDT>=? and FREPORTDT<=? order by FREPORTDT"
    Set req1_qr = rdb.CreateQuery("reqdtm1", str)
    
    ''SELECT QUERY FOR RREQUISITN TABLE
    str = "select * from rislib.rrequisitn where reqnno=?"
    Set req2_qr = rdb.CreateQuery("reqdtm2", str)
  
    ''SELECT QUERY FOR DTMVIEW1 TABLE
    str = "select * from dislib.DTMVIEW1 where FREPORTDT BETWEEN ? AND ? AND REQNSTAT='E' and MODALITYCD='SS' ORDER BY CASENO"
    Set DrequisitnQR = rdb.CreateQuery("DREQUISITN1", str)
    
    ''SELECT QUERY FOR DTMVIEW1 TABLE
    str = "select * from dislib.DTMVIEW1 where FREPORTDT BETWEEN ? AND ? AND REPTTM BETWEEN ? AND ? ORDER BY caseno"
    Set Drequisitn_QR = rdb.CreateQuery("DREQUISITN1111", str)
    
    ''*******AWAITED CASES
    ''SELECT QUERY FOR DTMTAB AND DTMVIEW1 TABLE
    str = "select DTMTAB.*,DTMVIEW1.* from DISLIB.DTMTAB, dislib.DTMVIEW1 where DTMVIEW1.reqno=dtmtab.reqno and DTMVIEW1.FREPORTDT BETWEEN ? AND ? AND DTMVIEW1.REQNSTAT='T' and DTMVIEW1.MODALITYCD='SS' and (DTMTAB.rh in('Awaited','Repeat') or DTMTAB.abo in('Awaited','Repeat') or DTMTAB.abs in('Awaited','Repeat') or DTMTAB.dat in('Awaited','Repeat') or DTMTAB.iat in('Awaited','Repeat')) ORDER BY DTMVIEW1.FREPORTDT DESC" 'or DTMTAB.hbsag in('Awaited','Repeat') or DTMTAB.hcv in('Awaited','Repeat') or DTMTAB.hiv in('Awaited','Repeat')) ORDER BY DTMVIEW1.FREPORTDT DESC"
    Set Drequisitn_QR_RCS = rdb.CreateQuery("DREQUISITN_QR_RCS", str)
    
    ''SELECT QUERY FOR DTMTAB AND DTMVIEW1 TABLE
    str = "select DTMTAB.*,DTMVIEW1.* from DISLIB.DTMTAB, dislib.DTMVIEW1 where DTMVIEW1.reqno=dtmtab.reqno and DTMVIEW1.FREPORTDT BETWEEN ? AND ? AND DTMVIEW1.REQNSTAT='T' and DTMVIEW1.MODALITYCD='SS' and (DTMTAB.HBSAG in('Awaited','Repeat')) ORDER BY DTMVIEW1.FREPORTDT DESC" 'or DTMTAB.hbsag in('Awaited','Repeat') or DTMTAB.hcv in('Awaited','Repeat') or DTMTAB.hiv in('Awaited','Repeat')) ORDER BY DTMVIEW1.FREPORTDT DESC"
    Set Drequisitn_QR_HBSAG = rdb.CreateQuery("DREQUISITN_QR_HBSAG", str)
    
    ''SELECT QUERY FOR DTMTAB AND DTMVIEW1 TABLE
    str = "select DTMTAB.*,DTMVIEW1.* from DISLIB.DTMTAB, dislib.DTMVIEW1 where DTMVIEW1.reqno=dtmtab.reqno and DTMVIEW1.FREPORTDT BETWEEN ? AND ? AND DTMVIEW1.REQNSTAT='T' and DTMVIEW1.MODALITYCD='SS' and (DTMTAB.HCV in('Awaited','Repeat')) ORDER BY DTMVIEW1.FREPORTDT DESC" 'or DTMTAB.hbsag in('Awaited','Repeat') or DTMTAB.hcv in('Awaited','Repeat') or DTMTAB.hiv in('Awaited','Repeat')) ORDER BY DTMVIEW1.FREPORTDT DESC"
    Set Drequisitn_QR_HCV = rdb.CreateQuery("DREQUISITN_QR_HCV", str)
    
    ''SELECT QUERY FOR DTMTAB AND DTMVIEW1 TABLE
    str = "select DTMTAB.*,DTMVIEW1.* from DISLIB.DTMTAB, dislib.DTMVIEW1 where DTMVIEW1.reqno=dtmtab.reqno and DTMVIEW1.FREPORTDT BETWEEN ? AND ? AND DTMVIEW1.REQNSTAT='T' and DTMVIEW1.MODALITYCD='SS' and (DTMTAB.HIV in('Awaited','Repeat')) ORDER BY DTMVIEW1.FREPORTDT DESC" 'or DTMTAB.hbsag in('Awaited','Repeat') or DTMTAB.hcv in('Awaited','Repeat') or DTMTAB.hiv in('Awaited','Repeat')) ORDER BY DTMVIEW1.FREPORTDT DESC"
    Set Drequisitn_QR_HIV = rdb.CreateQuery("DREQUISITN_QR_HIV", str)
    
    ''*******AWAITED CASES
    
    ''SELECT QUERY FOR DTMVIEW1 TABLE
    'str = "select * from dislib.DTMVIEW1 where reqno=?"
    'Commented and added by Shraddha on 21/04/10
    str = "select * from dislib.DTMVIEW where reqno=?"
    Set selreQR = rdb.CreateQuery("dtmre", str)
  
    ''SELECT QUERY FOR DTMTAB TABLE
    '******* Distinct clause added by Shraddha Pandit on 18/02/09 ********
    str = "select Distinct(REQNO),CASENO,ABO,RH,ABS,DAT,IAT,HBSAG,HCV,HIV,DTMLAB,REPTTM,HBSAGKEY,HCVKEY,HIVKEY FROM dislib.dtmtab WHERE reqno=? and caseno=? " ' not in('Awaited','Repeat') OR abo not in('Awaited','Repeat') OR abs not in('Awaited','Repeat') OR dat not in('Awaited','Repeat') OR iat not in('Awaited','Repeat') OR hbsag not in('Awaited','Repeat') OR hcv not in('Awaited','Repeat') OR hiv not in('Awaited','Repeat')"
    Set selDTMTabQR = rdb.CreateQuery("selDTMTabSTR", str)
    
    ''SELECT QUERY FOR DTMTAB TABLE
    str = "select * FROM dislib.dtmtab WHERE reqno=? and caseno=? and (rh in('Awaited','Repaet') or abo in('Awaited','Repaet') or abs in('Awaited','Repaet') or dat in('Awaited','Repaet') or iat in('Awaited','Repaet') or hbsag in('Awaited','Repaet') or hcv in('Awaited','Repaet') or hiv in('Awaited','Repaet'))"
    Set selDTMTabQR2 = rdb.CreateQuery("selDTMTabSTR2", str)
  
    ''SELECT QUERY FOR DTMTAB TABLE
    'str = "select DISTINCT(REQNO),CASENO,ABO,RH,ABS,DAT,IAT,HBSAG,HCV,HIV,DTMLAB,REPTTM,HBSAGKEY,HCVKEY,HIVKEY FROM dislib.dtmtab WHERE reqno=?"
    'Commented and addded by Shraddha on 28/04/10
    'str = "select DISTINCT(REQNO),CASENO,ABO,RH,ABS,DAT,IAT,HBSAG,HCV,HIV,DTMLAB,REPTTM,HBSAGKEY,HCVKEY,HIVKEY FROM dislib.dtmtab WHERE reqno=?"
    str = "select * FROM dislib.dtmtab WHERE REQNO=? and (ABO<>'' or RH <>'' or ABS<>'' or DAT<>'' or IAT<>'')"
    Set selDTMTabQR1 = rdb.CreateQuery("selDTMTabSTR111", str)
    
    ''SELECT QUERY FOR DTMTAB TABLE
    str = "select * from dislib.dTMtab where reqno=? AND CASENO=?"
    Set finddtmQR = rdb.CreateQuery("finddtmSTR", str)
    
    ''SELECT QUERY FOR DTMTAB TABLE
    str = "insert into dislib.dtmtab(caseno,REQNO,abo,rh,abs,dat,iat,hbsag,hcv,hiv,dtmlab,repttm,hbsagkey,hcvkey,hivkey) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set INSDTMTABQR = rdb.CreateQuery("INSDTMTABSTR", str)
    
    '' INSERT QUERY FOR DTMNO TABLE 'AG
    'STR = "insert into dislib.dtmno(CASENO,DTMNO,FLD1) values(?,?,?)"
    'Set INSDTMNOQR = rdb.CreateQuery("INSDTMNO_STR", STR)
    
    
    ''UPDATE QUERY FOR RREQUISITN TABLE
    str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=? WHERE REQNNO=?"
    Set UPDREQQR = rdb.CreateQuery("UPDREQSTR", str)
    
    ''UPDATE QUERY FOR RREQUISITN TABLE for prelim
    str = "UPDATE RISLIB.RREQUISITN SET PREPORTBY=?, PREPORTDT=?, REQNSTAT=? WHERE REQNNO=?"
    Set UPDREQQR_P = rdb.CreateQuery("UPDprel", str)
    
    
    ''UPDATE QUERY FOR RREQUISITN TABLE
    str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?,REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
    Set CUPDREQQR = rdb.CreateQuery("CUPDREQSTR", str)
    
    'INSERT QUERY FOR DMBCCOMM TABLE
    str = "INSERT INTO DISLIB.DMBCCOMM(REQNO,COMMENTS) VALUES(?,?)"
    Set INSDMBCCOMMQR = rdb.CreateQuery("INSDMBCCOMMQRDTM", str)

    'UPDATE QUERY FOR DMBCCOMM TABLE
    str = "UPDATE DISLIB.DMBCCOMM SET COMMENTS=? WHERE REQNO=?"
    Set UPDMBCCOMMQR = rdb.CreateQuery("UPDDMBCCOMMSTR", str)

    ''SELECT QUERY FOR DTMVIEW1 TABLE
    str = "SELECT * FROM DISLIB.DTMVIEW1 WHERE CASENO=? AND MODALITYCD=? AND (REQNSTAT=? OR REQNSTAT=?) ORDER BY FREPORTDT DESC "
    Set SELREQUIQR = rdb.CreateQuery("SELREQUISTR", str)
    
    ''SELECT QUERY FOR RREQUISITN TABLE
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=?"
    Set VARREQQR = rdb.CreateQuery("VARREQSTR", str)
        
    ''SELECT QUERY FOR DTMTAB TABLE
    str = "SELECT DTMTAB.REQNO FROM DISLIB.DTMTAB,DISLIB.DTMVIEW1 WHERE DTMVIEW1.REQNO=DTMTAB.REQNO AND DTMVIEW1.REQNODT>=? AND DTMVIEW1.REQNODT<=? "
    Set DTMVIEW1ALLQR = rdb.CreateQuery("DTMVIEW1ALLSTR", str)
    
    ''SELECT QUERY FOR TA101 TABLE
    str = "select * from PATOOLSLIB.ta101 where case_no =?"
    Set SEL_detail_qr = rdb.CreateQuery("select13", str)

    ''SELECT QUERY FOR DTMUSER TABLE
    str = "select * from dislib.dtmuser"
    Set DTMUSER_QR = rdb.CreateQuery("DTMUSER_STR", str)
    
    ''SELECT QUERY FOR DTMUSER TABLE
    str = "SELECT * FROM DISLIB.DTMUSER WHERE USERNAME=?"
    Set SELDTMUSER_QR = rdb.CreateQuery("SELDTMUSER_STR1", str)
       
    ''SELECT QUERY FOR EMPBASINFO TABLE
    str = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
    Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO_QR111101", str)
    
    '' SELECT QUERY FOR DREQSTAFFD TABLE
    str = "SELECT PNAME FROM DISLIB.DREQSTAFFD WHERE EMPCODE=? AND REQNO=?" 'H1
    Set SELEDREQSTAFFD_QR = rdb.CreateQuery("SELE11", str)
    ''*******BLANK CASES
    ''SELECT QUERY FOR DTMTAB AND DTMVIEW1 TABLE
    str = "select DTMTAB.*,DTMVIEW1.* from DISLIB.DTMTAB, dislib.DTMVIEW1 where DTMVIEW1.reqno=dtmtab.reqno and DTMVIEW1.FREPORTDT BETWEEN ? AND ? AND DTMVIEW1.REQNSTAT='T' and DTMVIEW1.MODALITYCD='SS' and (DTMTAB.rh IS NULL or TRIM(RH)='') AND (DTMTAB.abo IS NULL or TRIM(ABO)='') AND (DTMTAB.abs IS NULL or TRIM(ABS)='') AND (DTMTAB.dat IS NULL or TRIM(DAT)='') AND (DTMTAB.iat IS NULL or TRIM(IAT)='') ORDER BY DTMVIEW1.FREPORTDT DESC" 'or DTMTAB.hbsag in('Awaited','Repeat') or DTMTAB.hcv in('Awaited','Repeat') or DTMTAB.hiv in('Awaited','Repeat')) ORDER BY DTMVIEW1.FREPORTDT DESC"
    Set Drequisitn_QR_RCS_BLANK = rdb.CreateQuery("DREQUISITN_QR_RCS_BLANK", str)
    
    ''SELECT QUERY FOR DTMTAB AND DTMVIEW1 TABLE
    str = "select DTMTAB.*,DTMVIEW1.* from DISLIB.DTMTAB, dislib.DTMVIEW1 where DTMVIEW1.reqno=dtmtab.reqno and DTMVIEW1.FREPORTDT BETWEEN ? AND ? AND DTMVIEW1.REQNSTAT='T' and DTMVIEW1.MODALITYCD='SS' and DTMTAB.HBSAG IS NULL ORDER BY DTMVIEW1.FREPORTDT DESC" 'or DTMTAB.hbsag in('Awaited','Repeat') or DTMTAB.hcv in('Awaited','Repeat') or DTMTAB.hiv in('Awaited','Repeat')) ORDER BY DTMVIEW1.FREPORTDT DESC"
    Set Drequisitn_QR_HBSAG_BLANK = rdb.CreateQuery("DREQUISITN_QR_HBSAG_BLANK", str)
    
    ''SELECT QUERY FOR DTMTAB AND DTMVIEW1 TABLE
    str = "select DTMTAB.*,DTMVIEW1.* from DISLIB.DTMTAB, dislib.DTMVIEW1 where DTMVIEW1.reqno=dtmtab.reqno and DTMVIEW1.FREPORTDT BETWEEN ? AND ? AND DTMVIEW1.REQNSTAT='T' and DTMVIEW1.MODALITYCD='SS' and DTMTAB.HCV IS NULL ORDER BY DTMVIEW1.FREPORTDT DESC" 'or DTMTAB.hbsag in('Awaited','Repeat') or DTMTAB.hcv in('Awaited','Repeat') or DTMTAB.hiv in('Awaited','Repeat')) ORDER BY DTMVIEW1.FREPORTDT DESC"
    Set Drequisitn_QR_HCV_BLANK = rdb.CreateQuery("DREQUISITN_QR_HCV_BLANK", str)
    
    ''SELECT QUERY FOR DTMTAB AND DTMVIEW1 TABLE
    str = "select DTMTAB.*,DTMVIEW1.* from DISLIB.DTMTAB, dislib.DTMVIEW1 where DTMVIEW1.reqno=dtmtab.reqno and DTMVIEW1.FREPORTDT BETWEEN ? AND ? AND DTMVIEW1.REQNSTAT='T' and DTMVIEW1.MODALITYCD='SS' and DTMTAB.HIV IS NULL ORDER BY DTMVIEW1.FREPORTDT DESC" 'or DTMTAB.hbsag in('Awaited','Repeat') or DTMTAB.hcv in('Awaited','Repeat') or DTMTAB.hiv in('Awaited','Repeat')) ORDER BY DTMVIEW1.FREPORTDT DESC"
    Set Drequisitn_QR_HIV_BLANK = rdb.CreateQuery("DREQUISITN_QR_HIV_BLANK", str)
    
    ''*******BLANK CASES
    
    
    
End If
    MSFrcs.col = 3
    cmb_abo.Width = MSFrcs.CellWidth
    cmb_abo.Left = MSFrcs.CellLeft + MSFrcs.Left
    cmb_abo.Top = MSFrcs.CellTop + MSFrcs.Top
    
    MSFrcs.col = 4
    cmb_rh.Width = MSFrcs.CellWidth
    cmb_rh.Left = MSFrcs.CellLeft + MSFrcs.Left
    cmb_rh.Top = MSFrcs.CellTop + MSFrcs.Top
    
    MSFrcs.col = 5
    cmb_abs.Width = MSFrcs.CellWidth
    cmb_abs.Left = MSFrcs.CellLeft + MSFrcs.Left
    cmb_abs.Top = MSFrcs.CellTop + MSFrcs.Top
  
    MSFrcs.col = 6
    cmb_dat.Width = MSFrcs.CellWidth
    cmb_dat.Left = MSFrcs.CellLeft + MSFrcs.Left
    cmb_dat.Top = MSFrcs.CellTop + MSFrcs.Top
    
    MSFrcs.col = 7
    cmb_iat.Width = MSFrcs.CellWidth
    cmb_iat.Left = MSFrcs.CellLeft + MSFrcs.Left
    cmb_iat.Top = MSFrcs.CellTop + MSFrcs.Top
  
  ''FILL DATA IN OPTION COMBOBOX
  'cmb_enterby.List (0)
  'cmb_option.List(0) = "RCS"
  cmb_option.AddItem "RCS"
  'cmb_option.Text = "RCS" '28/04/10
  ' commented as all Three are shifted to HBSAG(ELISA),HCV(ELISA),HIV(ELISA)
  'cmb_option.AddItem "HBsAg(ELISA)"
  'cmb_option.AddItem "HCV(ELISA)"
  'cmb_option.AddItem "HIV(ELISA)"
  
  SSTab1.Enabled = False
  
  
  cmb_abo.Visible = False
  cmb_abs.Visible = False
  cmb_rh.Visible = False
  cmb_dat.Visible = False
  cmb_iat.Visible = False
  cmb_hbsag.Visible = False
  cmb_anti_hiv.Visible = False
  cmb_anti_hcv.Visible = False
    
  ''FILL DATA IN ABO COMBOBOX
    cmb_abo.AddItem "A"
    cmb_abo.AddItem "B"
    cmb_abo.AddItem "AB"
    cmb_abo.AddItem "O"
    cmb_abo.AddItem "Awaited"
    cmb_abo.AddItem "Repeat"
    cmb_abo.AddItem " "
    cmb_abo.AddItem "Bombay"
    
    ''FILL DATA IN RH COMBOBOX
    cmb_rh.AddItem "Negative"
    cmb_rh.AddItem "Positive"
    cmb_rh.AddItem "Weak D"
    cmb_rh.AddItem "Awaited"
    cmb_rh.AddItem "Repeat"
    cmb_rh.AddItem " "
    
    ''FILL DATA IN ANTI_HCV COMBOBOX
    cmb_anti_hcv.AddItem "Indeterminate"
    cmb_anti_hcv.AddItem "Non Reactive"
    cmb_anti_hcv.AddItem "Reactive"
    cmb_anti_hcv.AddItem "Awaited"
    cmb_anti_hcv.AddItem "Repeat"
    cmb_anti_hcv.AddItem " "
    
    ''FILL DATA IN DAT COMBOBOX
    cmb_dat.AddItem "Negative"
    cmb_dat.AddItem "Positive"
    cmb_dat.AddItem "Awaited"
    cmb_dat.AddItem "Repeat"
    cmb_dat.AddItem " "
    
    ''FILL DATA IN IAT COMBOBOX
    cmb_iat.AddItem "Negative"
    cmb_iat.AddItem "Positive"
    cmb_iat.AddItem "Awaited"
    cmb_iat.AddItem "Repeat"
    cmb_iat.AddItem " "
    
'    ''FILL DATA IN ABS COMBOBOX
'    cmb_abs.AddItem "Negative"
'    cmb_abs.AddItem "Positive"
'    cmb_abs.AddItem "Awaited"
'    cmb_abs.AddItem "Repeat"
'    cmb_abs.AddItem " "
    
    
    ''FILL DATA IN ABS COMBOBOX
    cmb_abs.AddItem "Absent"
    cmb_abs.AddItem "Present"
    cmb_abs.AddItem "Awaited"
    cmb_abs.AddItem "Repeat"
    cmb_abs.AddItem " "
    
    
    ''FILL DATA IN HIV COMBOBOX
    cmb_anti_hiv.AddItem "Indeterminate"
    cmb_anti_hiv.AddItem "Non Reactive"
    cmb_anti_hiv.AddItem "Reactive"
    cmb_anti_hiv.AddItem "Awaited"
    cmb_anti_hiv.AddItem "Repeat"
    cmb_anti_hiv.AddItem " "
    
    ''FILL DATA IN HBSAG COMBOBOX
    cmb_hbsag.AddItem "Indeterminate"
    cmb_hbsag.AddItem "Non Reactive"
    cmb_hbsag.AddItem "Reactive"
    cmb_hbsag.AddItem "Awaited"
    cmb_hbsag.AddItem "Repeat"
    cmb_hbsag.AddItem " "
    
     'SUNDEEP dR.NARAYAN
    SSTab1.TabVisible(0) = True
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    'SUNDEEP dR.NARAYAN
    
    '***** Shraddha Pandit 31/08/09 *****
    'Commented By Shraddha on 28/04/10 since screen is of NZZ SS only.
'    Set sel_departmentrs = SELMA106_QR.OpenResultset(1, 3)
'    If Not sel_departmentrs.EOF Then
'        com_dept.clear
'        com_dept.AddItem sel_departmentrs("dept_name")
'        com_dept.Text = Trim(sel_departmentrs("dept_name"))
'    End If
'
'    Set sel_modalityrs = SELRMODALITY_QR.OpenResultset(1, 3)
'    If Not sel_modalityrs.EOF Then
'        com_modality.clear
'        Do While Not sel_modalityrs.EOF
'            com_modality.AddItem Trim(sel_modalityrs(1)) & "$" & Trim(sel_modalityrs(0))
'            com_modality.Text = Trim(sel_modalityrs(1)) & "$" & Trim(sel_modalityrs(0))
'            sel_modalityrs.MoveNext
'        Loop
'    End If
    '**************************************
    
    
'    Set DTMUSER_RS = DTMUSER_QR.OpenResultset(1, 3)
'    If Not DTMUSER_RS.EOF Then
'        Do While Not DTMUSER_RS.EOF
'            txtusername.AddItem DTMUSER_RS("USERNAME") ''FILL DATA IN USERNAME COMBOBOX
'            cmb_enterby.AddItem DTMUSER_RS("USERNAME") ''FILL DATA IN ENTERBY COMBOBOX
'            DTMUSER_RS.MoveNext
'        Loop
'    End If
'
Exit Sub
MSG:
    MsgBox Err.Description, vbCritical, "DTM lab"
    Exit Sub
End Sub


Private Sub Form_Unload(cancel As Integer)
usr = ""
End Sub



Private Sub MSFhbsAg_Click()
MSFhbsAg.col = 3
    cmb_hbsag.Width = MSFhbsAg.CellWidth
    cmb_hbsag.Left = MSFhbsAg.CellLeft + MSFhbsAg.Left
    cmb_hbsag.Top = MSFhbsAg.CellTop + MSFhbsAg.Top
End Sub

Private Sub MSFhcv_Click()
MSFhcv.col = 3
    cmb_anti_hcv.Width = MSFhcv.CellWidth
    cmb_anti_hcv.Left = MSFhcv.CellLeft + MSFhcv.Left
    cmb_anti_hcv.Top = MSFhcv.CellTop + MSFhcv.Top
End Sub

Private Sub MSFhiv_Click()
MSFhiv.col = 3
    cmb_anti_hiv.Width = MSFhiv.CellWidth
    cmb_anti_hiv.Left = MSFhiv.CellLeft + MSFhiv.Left
    cmb_anti_hiv.Top = MSFhiv.CellTop + MSFhiv.Top
End Sub

Private Sub MSFrcs_Click()

MSFrcs.col = 3
    cmb_abo.Width = MSFrcs.CellWidth
    cmb_abo.Left = MSFrcs.CellLeft + MSFrcs.Left
    cmb_abo.Top = MSFrcs.CellTop + MSFrcs.Top
    
    MSFrcs.col = 4
    cmb_rh.Width = MSFrcs.CellWidth
    cmb_rh.Left = MSFrcs.CellLeft + MSFrcs.Left
    cmb_rh.Top = MSFrcs.CellTop + MSFrcs.Top
    
    MSFrcs.col = 5
    cmb_abs.Width = MSFrcs.CellWidth
    cmb_abs.Left = MSFrcs.CellLeft + MSFrcs.Left
    cmb_abs.Top = MSFrcs.CellTop + MSFrcs.Top
  
    MSFrcs.col = 6
    cmb_dat.Width = MSFrcs.CellWidth
    cmb_dat.Left = MSFrcs.CellLeft + MSFrcs.Left
    cmb_dat.Top = MSFrcs.CellTop + MSFrcs.Top
    
    MSFrcs.col = 7
    cmb_iat.Width = MSFrcs.CellWidth
    cmb_iat.Left = MSFrcs.CellLeft + MSFrcs.Left
    cmb_iat.Top = MSFrcs.CellTop + MSFrcs.Top
    
  MSFrcs.col = 9
    TXT_COMM.Visible = True
    TXT_COMM.Width = MSFrcs.CellWidth
    TXT_COMM.Left = MSFrcs.CellLeft + MSFrcs.Left
    TXT_COMM.Top = MSFrcs.CellTop + MSFrcs.Top
    
    MSFrcs.col = 10
    chk_Commit.Visible = True
    chk_Commit.Width = MSFrcs.CellWidth
    chk_Commit.Left = MSFrcs.CellLeft + MSFrcs.Left
    chk_Commit.Top = MSFrcs.CellTop + MSFrcs.Top
    
'TXT_COMM.SetFocus
  If MSFrcs.col = 1 Then
  
  
  End If
  
  
End Sub

Private Sub msk_caseno_LostFocus()
Dim CASENO As String
Dim cno_pad As String
 'On Error GoTo MSG:
If Trim(msk_caseno.Text) = "" Then
    Exit Sub
End If


i = InStr(com_modality.Text, "$")
MOD_CODE = Mid(com_modality.Text, i + 1, 2)

If Opt_patient.Value = True Then
        CASENO = msk_caseno.Text
      ''  MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))   ''02/12/2022
      '''' Commented PADZERO FUNCTION INSTEAD USED NEWNOPAD FUNCTION
'cno_pad = padzero(5, "C", msk_caseno.Text, 3)
    
  If IsNumeric(Left(Trim(msk_caseno.Text), 1)) Then
        Dim Ks As Integer
        Dim str1s As String
        Dim strs As String
            strs = Trim(msk_caseno.Text)
            Ks = Len(strs)
            str1s = Mid(strs, 5, Ks)
              Dim sel_CASE_rs  As rdo.rdoResultset
              Set sel_CASE_rs = rdb.OpenResultset(" SELECT * FROM PATOOLSLIB.NEWCASETBL WHERE CASE_CODE = '" & Left(UCase(msk_caseno.FormattedText), 4) & "' ", 1, 1)
              If Not sel_CASE_rs.EOF Then
                    Dim CSNO As String
                    CSNO = Trim(sel_CASE_rs(1))
                    CSNO = CSNO & str1s
                     cno_pad = NEWNOpad(CSNO)
              Else
               cno_pad = NEWNOpad(UCase(msk_caseno))
              End If
  Else
    cno_pad = NEWNOpad(UCase(msk_caseno))
  End If
     msk_caseno.Text = cno_pad

    If SEL_detail_flag = False Then
            SEL_detail_qr(0) = msk_caseno.FormattedText
            Set SEL_detail_rs = SEL_detail_qr.OpenResultset(1, 3)
            SEL_detail_flag = True
        Else
            SEL_detail_qr(0) = msk_caseno.FormattedText
            SEL_detail_rs.MoreResults
            SEL_detail_rs.Requery
        End If
        If Not SEL_detail_rs.EOF Then
            txtcasename.Text = Trim(SEL_detail_rs("first_name")) & " " & Trim(SEL_detail_rs("mid_name")) & " " & Trim(SEL_detail_rs("last_name"))
        End If
    If NEWFLAG = True Then
          If REQFLAG1 = False Then
              req_qr(0) = msk_caseno.FormattedText
              req_qr(1) = Trim(MOD_CODE)
              Set req_rs = req_qr.OpenResultset(1, 3)
              REQFLAG1 = True
          Else
              req_qr(0) = msk_caseno.FormattedText
              req_qr(1) = Trim(MOD_CODE)
              req_rs.MoreResults
              req_rs.Requery
          End If
          If req_rs.RowCount > 0 Then
              Com_REQNO.clear
              Do While Not req_rs.EOF
                Com_REQNO.AddItem req_chglen(req_rs(0))
                req_rs.MoveNext
              Loop
          Else
            MsgBox " No Data Found !!", vbInformation, "DTM LAB"
          End If
    End If
    ''COMMENT
    If viewflag = True Then
        If SELREQUIFLAG = False Then
            SELREQUIQR(0) = msk_caseno.FormattedText
            SELREQUIQR(1) = Trim(MOD_CODE)
            'SELREQUIQR(1) = Trim("NZZ")
            SELREQUIQR(2) = "T"
            SELREQUIQR(3) = "F"
            Set SELREQUIRS = SELREQUIQR.OpenResultset(1, 3)
            SELREQUIFLAG = True
        Else
            SELREQUIQR(0) = msk_caseno.FormattedText
            SELREQUIQR(1) = Trim(MOD_CODE)
            'SELREQUIQR(1) = Trim("NZZ")
            SELREQUIQR(2) = "T"
            SELREQUIQR(3) = "F"
            SELREQUIRS.MoreResults
            SELREQUIRS.Requery
        End If
        If SELREQUIRS.RowCount > 0 Then
            Com_REQNO.clear
            While Not SELREQUIRS.EOF
                Com_REQNO.AddItem req_chglen(SELREQUIRS("REQNO"))
                SELREQUIRS.MoveNext
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
        If REQFLAG2 = False Then
            req_qr(0) = msk_caseno.FormattedText
            req_qr(1) = Trim(MOD_CODE)
            Set req_rs = req_qr.OpenResultset(1, 3)
            REQFLAG2 = True
        Else
            req_qr(0) = msk_caseno.FormattedText
            req_qr(1) = Trim(MOD_CODE)
            req_rs.MoreResults
            req_rs.Requery
        End If
        If Not req_rs.EOF Then
            Com_REQNO.clear
            Do While Not req_rs.EOF
                Com_REQNO.AddItem req_chglen(req_rs("REQNNO"))
                req_rs.MoveNext
            Loop
        End If
        
  'JUSTIN
   End If
        'If SELEMPBASINFOFLAG = False Then
         '   SELEMPBASINFO_QR(0) = msk_caseno.Text
          '  Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
           ' SELEMPBASINFOFLAG = True
        'Else
         '   SELEMPBASINFO_QR(0) = msk_caseno.Text
          '  SELEMPBASINFO_RS.MoreResults
           ' SELEMPBASINFO_RS.Requery
        'End If
    ' HIDE BY SANDEEP
'        If SELEDREQSTAFFD_FLAG = False Then
'            SELEDREQSTAFFD_QR(0) = msk_caseno.Text
'            SELEDREQSTAFFD_QR(1) = Com_REQNO.Text
'            Set SELEDREQSTAFFD_RS = SELEDREQSTAFFD_QR.OpenResultset(1, 3)
'            SELEDREQSTAFFD_FLAG = True
'        Else
'            SELEDREQSTAFFD_RS.MoreResults
'            SELEDREQSTAFFD_RS.Requery
'        End If

        'If SELEMPBASINFO_RS.RowCount > 0 Then
         '   txtcasename.Text = Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME"))
        'End If
' HIDE BY SANDEEP
'        If SELEDREQSTAFFD_RS.RowCount > 0 Then
'            txtcasename.Text = SELEDREQSTAFFD_RS("PNAME")
'        End If
    If viewflag = True Then
        If SELREQUIFLAG = False Then
            SELREQUIQR(0) = msk_caseno.FormattedText
            SELREQUIQR(1) = Trim(MOD_CODE)
            'SELREQUIQR(1) = Trim("NZZ")
            SELREQUIQR(2) = "T"
            SELREQUIQR(3) = "F"
            Set SELREQUIRS = SELREQUIQR.OpenResultset(1, 3)
            SELREQUIFLAG = True
        Else
            SELREQUIQR(0) = msk_caseno.FormattedText
            SELREQUIQR(1) = Trim(MOD_CODE)
            'SELREQUIQR(1) = Trim("NZZ")
            SELREQUIQR(2) = "T"
            SELREQUIQR(3) = "F"
            SELREQUIRS.MoreResults
            SELREQUIRS.Requery
        End If
        If SELREQUIRS.RowCount > 0 Then
            Com_REQNO.clear
            While Not SELREQUIRS.EOF
                Com_REQNO.AddItem req_chglen(SELREQUIRS("REQNO"))
                SELREQUIRS.MoveNext
            Wend
         Else
            MsgBox " No  Requisition for this Modality"
            btnenabl ("INI")
            Exit Sub
        End If
        If SELEDREQSTAFFD_FLAG = False Then
            SELEDREQSTAFFD_QR(0) = msk_caseno.Text
            SELEDREQSTAFFD_QR(1) = Com_REQNO.Text
            Set SELEDREQSTAFFD_RS = SELEDREQSTAFFD_QR.OpenResultset(1, 3)
            SELEDREQSTAFFD_FLAG = True
        Else
            SELEDREQSTAFFD_RS.MoreResults
            SELEDREQSTAFFD_RS.Requery
        End If
        'If SELEMPBASINFO_RS.RowCount > 0 Then
         '   txtcasename.Text = Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME"))
        'End If
        If SELEDREQSTAFFD_RS.RowCount > 0 Then
            txtcasename.Text = SELEDREQSTAFFD_RS("PNAME")
        End If
        
''        If SELEMPBASINFOFLAG = False Then
''            SELEMPBASINFO_QR(0) = msk_caseno.Text
''            Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
''            SELEMPBASINFOFLAG = True
''        Else
''            SELEMPBASINFO_QR(0) = msk_caseno.Text
''            SELEMPBASINFO_RS.MoreResults
''            SELEMPBASINFO_RS.Requery
''        End If
''        If SELEMPBASINFO_RS.RowCount > 0 Then
''            txtcasename.Text = Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME"))
''        End If
End If
    'End If
End If
Com_REQNO.SetFocus
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub Msk_fromdt_GotFocus()
If OPT_ALL.Value = True Then
    msk_fromdt.Text = Format(DATE11, "DD/MM/YYYY")
    msk_todt.Text = Format(DATE11, "DD/MM/YYYY")
ElseIf OPT_CASENO.Value = True Then
    msk_fromdt.Text = ""
    msk_todt.Text = ""
End If
End Sub

Private Sub msk_todt_LostFocus()
'If InStr(msk_todt.Text, "_") > 0 Then
'   msk_todt.Mask = ""
'            msk_todt.Text = ""
'            msk_todt.Mask = "##/##/####"
'            Exit Sub
'End If
'If IsDate(GET_REVDATE(msk_todt.Text)) = False Then
'    msk_todt.Mask = ""
'    msk_todt.Text = ""
'    msk_todt.Mask = "##/##/####"
'    MsgBox "Please ENTER the  correct DATE."
'    msk_todt.SetFocus
'    Exit Sub
'End If
'If msk_fromdt.Text <> "" And msk_todt.Text <> "" Then
'    If REQ1_QRFLAG = False Then
'        req1_qr(0) = msk_fromdt.FormattedText
'        req1_qr(1) = msk_todt.FormattedText
'        Set req1_rs = req1_qr.OpenResultset(1, 3)
'        REQ1_QRFLAG = True
'    Else
'        req1_qr(0) = msk_fromdt.FormattedText
'        req1_qr(1) = msk_todt.FormattedText
'        req1_rs.MoreResults
'        req1_rs.Requery
'    End If
'    If Not req1_rs.EOF Then
'        Do While Not req1_rs.EOF
'            Com_REQNO.AddItem req1_rs(0)
'            req1_rs.MoveNext
'        Loop
'    End If
'End If
'    'Com_RENO.SetFocus
'cmb_option.SetFocus
FTIME.SetFocus
End Sub


Private Sub Opt_all_LostFocus()
If OPT_ALL.Value = True Then
    Com_REQNO.Enabled = False
    msk_caseno.Enabled = False
    'OPT_CASENO.Enabled = False
    txtcasename.Enabled = False
    
    ''FILL DATA IN OPTION COMBOBOX
    cmb_option.clear
    cmb_option.AddItem "RCS"
    'cmb_option.Text = "RCS" 'Shraddha 28/04/10
    ' commented as all thee are shifted to Seriology  Sundeep
'''''    cmb_option.AddItem "HBsAg(ELISA)"
'''''    cmb_option.AddItem "HCV(ELISA)"
'''''    cmb_option.AddItem "HIV(ELISA)"
    'msk_fromdt.SetFocus
End If
End Sub

Private Sub OPT_CASENO_LostFocus()
If OPT_CASENO.Value = True Then
    If Opt_patient.Value = True Then
        'lbl_caseno.Caption = "Case No."
        'msk_caseno.Mask = "??>/#####"
        msk_caseno.Enabled = True
        msk_caseno.SetFocus
    ElseIf opt_staff.Value = True Then
        'lbl_caseno.Caption = "  CC No."
        'msk_caseno.Mask = "########"
        msk_caseno.Enabled = True
        msk_caseno.SetFocus
    End If
    'msk_caseno.SetFocus
    OPT_ALL.Enabled = False
    'cbo_modality.Enabled = False
    msk_caseno.Enabled = True
    Com_REQNO.Enabled = True
    msk_fromdt.Enabled = False
    msk_todt.Enabled = False
    msk_caseno.SetFocus
End If
End Sub

Private Sub Opt_patient_LostFocus()
If Opt_patient.Value = True Then
    'lbl_caseno.Caption = "Case No."
    'MaskEdBox1.Mask = "??>/#####"
    'MaskEdBox1.Enabled = True
    'MaskEdBox1.SetFocus
End If
End Sub

Private Sub opt_staff_LostFocus()
If opt_staff.Value = True Then

'    lbl_caseno.Caption = "  CC No."
'    MaskEdBox1.Mask = "########"
'    MaskEdBox1.Enabled = True
'    MaskEdBox1.SetFocus
End If
End Sub

Private Sub TTIME_LostFocus()
cmb_option.SetFocus
End Sub

Private Sub TXT_COMM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

MSFrcs.TextMatrix(MSFrcs.row, MSFrcs.col) = TXT_COMM.Text

  SEL_DMBCCOMM_QR(0) = req_chglen(MSFrcs.TextMatrix(MSFrcs.row, 1))
            Set SEL_DMBCCOMM_RS = SEL_DMBCCOMM_QR.OpenResultset(2, 1)
            If Not SEL_DMBCCOMM_RS.EOF Then
                UPD_DMBCCOMM_qr(0) = Trim(TXT_COMM.Text)
                UPD_DMBCCOMM_qr(1) = req_chglen(MSFrcs.TextMatrix(MSFrcs.row, 1))
                UPD_DMBCCOMM_qr.Execute
            Else
            INS_DMBCCOMM_QR(0) = req_chglen(MSFrcs.TextMatrix(MSFrcs.row, 1))
            INS_DMBCCOMM_QR(1) = Trim(TXT_COMM.Text)
            INS_DMBCCOMM_QR.Execute
            End If
                        
            MSFrcs.TextMatrix((MSFrcs.row), 9) = Trim(TXT_COMM.Text)
     '   End If
     TXT_COMM.Text = ""
'TXT_COMM.Visible = False
        TXT_COMM.Visible = False
End If
End Sub

Private Sub txtcasename_LostFocus()
'cmb_reqno.SetFocus
End Sub

Private Sub txtPassword_LostFocus()
'If SELDTMUSER_FLAG = False Then
'    SELDTMUSER_QR(0) = Trim(txtusername.Text)
'    Set SELDTMUSER_RS = SELDTMUSER_QR.OpenResultset(1, 3)
'    SELDTMUSER_FLAG = True
'Else
'    SELDTMUSER_QR(0) = Trim(txtusername.Text)
'    SELDTMUSER_RS.MoreResults
'    SELDTMUSER_RS.Requery
'End If
'If Not SELDTMUSER_RS.EOF Then
'    If Trim(SELDTMUSER_RS("password")) <> Trim(txtpassword.Text) Then
'        MsgBox " U r not an authorized person to log in !!", vbInformation, "DTM LAB"
'        txtpassword.Text = ""
'        txtusername.Text = ""
'        btnenabl ("INI")
'        CMD_NEW.SetFocus
'        'txtpassword.SetFocus
'        Exit Sub
'    Else
'        If sel_departmentflag = False Then
'               '

'(0) = "NZZ"
'            Set sel_departmentrs = SELMA106_QR.OpenResultset(1, 3)
'            sel_departmentflag = True
'        Else
'            'selma106_qr(0) = "NZZ"
'            sel_departmentrs.MoreResults
'            sel_departmentrs.Requery
'        End If
'        If Not sel_departmentrs.EOF Then
'            com_dept.clear
'            com_dept.AddItem sel_departmentrs("dept_name")
'        End If
'        If SEL_MODALITYFLAG1 = False Then
'            'SELRMODALITY_QR(0) = "NZZ"
'            Set sel_modalityrs = SELRMODALITY_QR.OpenResultset(1, 3)
'            SEL_MODALITYFLAG1 = True
'        Else
'            'SELRMODALITY_QR(0) = "NZZ"
'            sel_modalityrs.MoreResults
'            sel_modalityrs.Requery
'        End If
'        If Not sel_modalityrs.EOF Then
'            com_modality.clear
'            Do While Not sel_modalityrs.EOF
'                com_modality.AddItem Trim(sel_modalityrs(1)) & "$" & Trim(sel_modalityrs(0))
'                sel_modalityrs.MoveNext
'            Loop
'        End If
'        usr = txtusername.Text
'        dtmlab_report.Caption = "Dtm Lab : User Name : " & usr
'        txtusername.Text = ""
'        txtpassword.Text = ""
'        If NEWFLAG = True Then
'            cmb_enterby.Text = usr
'        End If
'    End If
'End If
End Sub

Private Sub txtusername_LostFocus()
'txtpassword.SetFocus
End Sub
Public Sub FLDENABL(EN As String)
If EN = "DIS" Then
        'cmb_dept.Enabled = False
        'cmb_modality.Enabled = False
        'MaskEdBox1.Enabled = False
        cmb_enterby.Enabled = False
End If
If EN = "ENA" Then
'        cmb_dept.Enabled = True
        'cmb_modality.Enabled = True
        'MaskEdBox1.Enabled = True
        'com_reqno.Enabled = True
        'COMMENTED BY SHWEYA 10-JUN-2013
        'cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
End If
If EN = "MOD" Then
        'cmb_dept.Enabled = True
        'cmb_modality.Enabled = True
        'MaskEdBox1.Enabled = True
        Com_REQNO.Enabled = True
        'TXT_PATH.Enabled = True
        'txt_name.Enabled = True
        'txt_sex.Enabled = True
        'txt_age.Enabled = True
        'txt_category.Enabled = True
        'txt_status.Enabled = True
        'txt_pridoc.Enabled = True
        'txt_ward.Enabled = True
        'txt_bedno.Enabled = True
       ' txt_modcode.Enabled = True
       ' txt_refdoc.Enabled = True
        'COMMENTED BY SHWETA 10-JUN-2013
       ' cmb_enterby.Enabled = True
        cmb_finalby.Enabled = True
End If
End Sub

Public Sub field_clear()
Dim obj As Object
chk_awaited.Value = 0
chk_blank.Value = 0
'com_dept.ListIndex = -1
com_dept.Text = ""
'com_modality.ListIndex = -1
com_modality.Text = ""
Com_REQNO.clear
Com_REQNO.ListIndex = -1
Opt_patient.Value = False
opt_staff.Value = False
If Opt_patient.Value = False Then
    msk_caseno.Mask = ""
    msk_caseno.Text = ""
   ' msk_caseno.Mask = ">??/#####"
End If
If opt_staff.Value = False Then
    msk_caseno.Mask = ""
    msk_caseno.Text = ""
   ' msk_caseno.Mask = "#######"
End If
txtcasename.Text = ""
msk_fromdt.Text = ""
msk_todt.Text = ""
usr = ""
dtmlab_report.Caption = "DTM Lab"
OPT_ALL.Enabled = True
OPT_CASENO.Enabled = True
cmb_option.Enabled = True
'cmb_enterby.Text = ""
cmb_finalby.Text = ""
cmb_option.ListIndex = -1
Com_REQNO.ListIndex = -1
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
'    cmb_abo.Visible = False
'    cmb_rh.Visible = False
'    cmb_abs.Visible = False
'    cmb_dat.Visible = False
'    cmb_iat.Visible = False
'    cmb_hbsag.Visible = False
'    cmb_anti_hcv.Visible = False
'    cmb_anti_hiv.Visible = False
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
End Sub
Public Sub TXT_CLEAR()
Dim obj As Object
For Each obj In dtmlab_report.Controls
If TypeOf obj Is TextBox Then
obj = ""
End If
Next
End Sub

Public Sub BCCLEAR()
'TXT_CLEAR
'cmb_dept.Text = ""
'cmb_modality.Text = ""
FTIME.Text = "##:##:##"
TTIME.Text = "##:##:##"
cmb_abo.Visible = False
cmb_rh.Visible = False
cmb_abs.Visible = False
cmb_dat.Visible = False
cmb_iat.Visible = False
MSFrcs.Rows = 1
MSFhbsAg.Rows = 1
MSFhcv.Rows = 1
MSFhiv.Rows = 1
cmb_hbsag.Visible = False
cmb_anti_hcv.Visible = False
cmb_anti_hiv.Visible = False
End Sub

Private Sub assign()
Dim NROWS
Dim HIV As String
Dim FLD1
Dim coment As String
Dim ctr As Integer
ctr = 0
'On Error GoTo MSG:
rdb.BeginTrans
If flagsave = 1 Then       'save
    If Trim(cmb_option.Text) = "RCS" Then
        i = 1
        For i = 1 To MSFrcs.Rows - 1
        
            If MSFrcs.TextMatrix(i, 1) <> "" Then
               'If (Len(Trim(MSFrcs.TextMatrix(i, 3)))) = 0 And (Len(Trim(MSFrcs.TextMatrix(i, 4)))) = 0 And (Len(Trim(MSFrcs.TextMatrix(i, 5)))) = 0 And (Len(Trim(MSFrcs.TextMatrix(i, 6)))) = 0 Then
               
                   ' ONLY IF ABO Rh AND DAT IAT ABO ARE ENTERED
            'Next if statement is added by Shraddha Pandit on 20/02/09
            'Insert in DISLIB/DTMTAB only if either of ABO/Rh/Abs/DAT/IAT is entered
            'If (MSFrcs.TextMatrix(MSFrcs.row, 3) = "") And (MSFrcs.TextMatrix(MSFrcs.row, 4) = "") And (MSFrcs.TextMatrix(I, 5) = "") And (MSFrcs.TextMatrix(I, 6) = "") And (MSFrcs.TextMatrix(I, 7) = "") Then
            If (MSFrcs.TextMatrix(i, 3) = "") And (MSFrcs.TextMatrix(i, 4) = "") And (MSFrcs.TextMatrix(i, 5) = "") And (MSFrcs.TextMatrix(i, 6) = "") And (MSFrcs.TextMatrix(i, 7) = "") Then
               ' MsgBox "Please enter ABO/Rh/Abs/DAT/IAT"
            Else
            
                INSDTMTABQR(0) = Trim(MSFrcs.TextMatrix(i, 2))  'CASENO
                INSDTMTABQR(1) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1))) 'REQNO
                INSDTMTABQR(2) = "" & Trim(MSFrcs.TextMatrix(i, 3)) 'ABO
                INSDTMTABQR(3) = "" & Trim(MSFrcs.TextMatrix(i, 4)) 'RH
                'FLD1 = Trim(MSFrcs.TextMatrix(i, 3)) & " " & Trim(MSFrcs.TextMatrix(i, 4))
                'INSDTMNOQR(0) = Trim(MSFrcs.TextMatrix(i, 2)) 'CASENO
                'INSDTMNOQR(1) = "" & Trim(MSFrcs.TextMatrix(i, 8)) 'DTMNO
                'INSDTMNOQR(2) = FLD1
                INSDTMTABQR(4) = "" & Trim(MSFrcs.TextMatrix(i, 5)) 'ABS
                INSDTMTABQR(5) = "" & Trim(MSFrcs.TextMatrix(i, 6)) 'DAT
                INSDTMTABQR(6) = "" & Trim(MSFrcs.TextMatrix(i, 7)) 'IAT
                INSDTMTABQR(7) = Null 'HBSAG
                INSDTMTABQR(8) = Null 'HCV
                INSDTMTABQR(9) = Null 'HIV
                INSDTMTABQR(10) = "" & Trim(MSFrcs.TextMatrix(i, 8)) 'DTMNO
                INSDTMTABQR(11) = time1 'TIME
                INSDTMTABQR(12) = 1 'HBSAGKEY
                INSDTMTABQR(13) = 0 'HCVKEY
                INSDTMTABQR(14) = 0 'HIVKEY
                INSDTMTABQR.Execute
                
                'HERE
                
                
                '''******UPDATE THE FLD1 IN DTMNO TABLE
                UPDDTMNO_QR(0) = Trim(MSFrcs.TextMatrix(i, 3)) & " " & Trim(MSFrcs.TextMatrix(i, 4))
                UPDDTMNO_QR(1) = Trim(MSFrcs.TextMatrix(i, 2))
                UPDDTMNO_QR.Execute
                
            End If
            
            'following txt_assigning is added by shweta 03-04-2013
            txt_comments.Text = Trim(MSFrcs.TextMatrix(i, 9))
            '''
            
                If SELDMBC_FLAG = False Then
                   SELDMBC_QR(0) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
                  Set SELDMBC_RS = SELDMBC_QR.OpenResultset(1, 3)
                 SELDMBC_FLAG = True
                Else
                   SELDMBC_QR(0) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
                   SELDMBC_RS.MoreResults
                  SELDMBC_RS.Requery
                End If
                If SELDMBC_RS.RowCount = 0 Then
                    INSDMBCCOMMQR(0) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
                    If Len(Trim(txt_comments.Text)) > 0 Then
                        INSDMBCCOMMQR(1) = Null
                        INSDMBCCOMMQR(1) = txt_comments.Text
                    Else
                        INSDMBCCOMMQR(1) = Null
                    End If
                    INSDMBCCOMMQR.Execute
                Else
                  If Len(Trim(txt_comments.Text)) > 0 Then
                      UPDMBCCOMMQR(0) = Null
                     UPDMBCCOMMQR(0) = txt_comments.Text
                 Else
                    UPDMBCCOMMQR(0) = Null
                 End If
                    UPDMBCCOMMQR(1) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
                        UPDMBCCOMMQR.Execute
               End If
txt_comments.Text = ""
                'End If
                'ADDED BY MILIND DT. 25/01/06 FOR DISPLAYING ONLY REQNNO OF STATUS 'E'
                'Commented by Shraddha on 28/04/10
                'if record is not beig inserted in DTMTAB then no need to update rrequisitn.
'                If (MSFrcs.TextMatrix(I, 3) = "") And (MSFrcs.TextMatrix(I, 4) = "") And (MSFrcs.TextMatrix(I, 5) = "") And (MSFrcs.TextMatrix(I, 6) = "") And (MSFrcs.TextMatrix(I, 7) = "") Then
'
'                    UPDREQQR(0) = Trim(cmb_enterby.Text)
'                    UPDREQQR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
'                    UPDREQQR(2) = "E"
'                    UPDREQQR(3) = Trim(req_chglen(MSFrcs.TextMatrix(I, 1)))
'                    UPDREQQR.Execute
'                Else
                    If (MSFrcs.TextMatrix(i, 3) <> "") Or (MSFrcs.TextMatrix(i, 4) <> "") Or (MSFrcs.TextMatrix(i, 5) <> "") Or (MSFrcs.TextMatrix(i, 6) <> "") Or (MSFrcs.TextMatrix(i, 7) <> "") Then
                        ctr = ctr + 1
                        UPDREQQR(0) = Trim(cmb_enterby.Text)
                        UPDREQQR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                        UPDREQQR(2) = "T"
                        UPDREQQR(3) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
                        UPDREQQR.Execute
                        
                        
                        UPDREQQR_P(0) = Trim(cmb_enterby.Text)
                UPDREQQR_P(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                UPDREQQR_P(2) = "T"
                UPDREQQR_P(3) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
                UPDREQQR_P.Execute
                        
                    End If
                'End If
                'UPTO HERE
                
            End If
         ' End If  ' ONLY IF ABO Rh AND DAT IAT ABO ARE ENTERED
          
        Next i
        'commented as all three HBSAG(ELISA) ,HCV,HIV are shifted to Seriology Sundeep
''''''    ElseIf Trim(cmb_option.Text) = "HBsAg(ELISA)" Then
''''''        i = 1
''''''        For i = 1 To MSFhbsAg.Rows - 1
''''''            If MSFhbsAg.TextMatrix(i, 1) <> "" Then
''''''                INSDTMTABQR(0) = Trim(MSFhbsAg.TextMatrix(i, 2)) 'CASENO
''''''                INSDTMTABQR(1) = Trim(req_chglen(MSFhbsAg.TextMatrix(i, 1))) 'REQNO
''''''                INSDTMTABQR(2) = Null 'ABO
''''''                INSDTMTABQR(3) = Null 'RH
''''''                INSDTMTABQR(4) = Null 'ABS
''''''                INSDTMTABQR(5) = Null 'DAT
''''''                INSDTMTABQR(6) = Null 'IAT
''''''                If Trim(MSFhbsAg.TextMatrix(i, 3)) = "" Then 'HBSAG
''''''                    INSDTMTABQR(7) = Null
''''''                Else
''''''                    a = makekey
''''''                    INSDTMTABQR(7) = transform(Trim(MSFhbsAg.TextMatrix(i, 3)), a)
''''''                End If
''''''                INSDTMTABQR(8) = Null 'HCV
''''''                INSDTMTABQR(9) = Null 'HIV
''''''                INSDTMTABQR(10) = "" & Trim(MSFhbsAg.TextMatrix(i, 4)) 'DTMNO
''''''                INSDTMTABQR(11) = time1 'TIME
''''''                INSDTMTABQR(12) = CInt(a) 'KEY
''''''                INSDTMTABQR(13) = 0 'KEY
''''''                INSDTMTABQR(14) = 0 'KEY
''''''                INSDTMTABQR.Execute
''''''
''''''                INSDMBCCOMMQR(0) = Trim(req_chglen(MSFhbsAg.TextMatrix(i, 1)))
''''''                If Len(Trim(txt_comments.Text)) > 0 Then
''''''                    INSDMBCCOMMQR(1) = Null
''''''                    INSDMBCCOMMQR(1) = txt_comments.Text
''''''                Else
''''''                    INSDMBCCOMMQR(1) = Null
''''''                End If
''''''                INSDMBCCOMMQR.Execute
''''''
''''''                UPDREQQR(0) = Trim(cmb_enterby.Text)
''''''                UPDREQQR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''''''                UPDREQQR(2) = "T"
''''''                UPDREQQR(3) = Trim(req_chglen(MSFhbsAg.TextMatrix(i, 1)))
''''''                UPDREQQR.Execute
''''''            End If
''''''        Next i
''''''    ElseIf Trim(cmb_option.Text) = "HCV(ELISA)" Then
''''''        i = 1
''''''        For i = 1 To MSFhcv.Rows - 1
''''''            If MSFhcv.TextMatrix(i, 1) <> "" Then
''''''                INSDTMTABQR(0) = Trim(MSFhcv.TextMatrix(i, 2)) 'CASENO
''''''                INSDTMTABQR(1) = Trim(req_chglen(MSFhcv.TextMatrix(i, 1))) 'REQNO
''''''                INSDTMTABQR(2) = Null 'ABO
''''''                INSDTMTABQR(3) = Null 'RH
''''''                INSDTMTABQR(4) = Null 'ABS
''''''                INSDTMTABQR(5) = Null 'DAT
''''''                INSDTMTABQR(6) = Null 'IAT
''''''                INSDTMTABQR(7) = Null 'HBSAG
''''''                If Trim(MSFhcv.TextMatrix(i, 3)) = "" Then 'HCV
''''''                    INSDTMTABQR(8) = Null
''''''                Else
''''''                    a = makekey
''''''                    INSDTMTABQR(8) = transform(Trim(MSFhcv.TextMatrix(i, 3)), a)
''''''                End If
''''''                INSDTMTABQR(9) = Null 'HIV
''''''                INSDTMTABQR(10) = "" & Trim(MSFhcv.TextMatrix(i, 4)) 'DTMNO
''''''                INSDTMTABQR(11) = time1 'TIME
''''''                INSDTMTABQR(12) = 0 'KEY
''''''                INSDTMTABQR(13) = CInt(a) 'KEY
''''''                INSDTMTABQR(14) = 0 'KEY
''''''                INSDTMTABQR.Execute
''''''
''''''                INSDMBCCOMMQR(0) = Trim(req_chglen(MSFhcv.TextMatrix(i, 1)))
''''''                If Len(Trim(txt_comments.Text)) > 0 Then
''''''                    INSDMBCCOMMQR(1) = Null
''''''                    INSDMBCCOMMQR(1) = txt_comments.Text
''''''                Else
''''''                    INSDMBCCOMMQR(1) = Null
''''''                End If
''''''                INSDMBCCOMMQR.Execute
''''''
''''''                UPDREQQR(0) = Trim(cmb_enterby.Text)
''''''                UPDREQQR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''''''                UPDREQQR(2) = "T"
''''''                UPDREQQR(3) = Trim(req_chglen(MSFhcv.TextMatrix(i, 1)))
''''''                UPDREQQR.Execute
''''''            End If
''''''        Next i
''''''    ElseIf Trim(cmb_option.Text) = "HIV(ELISA)" Then
''''''        i = 1
''''''        Dim RS As rdoResultset
''''''        For i = 1 To MSFhiv.Rows - 1
''''''            If MSFhiv.TextMatrix(i, 1) <> "" Then
''''''
''''''                INSDTMTABQR(0) = Trim(MSFhiv.TextMatrix(i, 2)) 'CASENO
''''''                INSDTMTABQR(1) = Trim(req_chglen(MSFhiv.TextMatrix(i, 1))) 'REQNO
''''''                INSDTMTABQR(2) = Null 'ABO
''''''                INSDTMTABQR(3) = Null 'RH
''''''                INSDTMTABQR(4) = Null 'ABS
''''''                INSDTMTABQR(5) = Null 'DAT
''''''                INSDTMTABQR(6) = Null 'IAT
''''''                INSDTMTABQR(7) = Null 'HSBAG
''''''                INSDTMTABQR(8) = Null 'HCV
''''''                If Trim(MSFhiv.TextMatrix(i, 3)) = "" Then 'HIV
''''''                    INSDTMTABQR(9) = Null
''''''                Else
''''''                    a = makekey
''''''                    INSDTMTABQR(9) = transform(Trim(MSFhiv.TextMatrix(i, 3)), a)
''''''                End If
''''''                INSDTMTABQR(10) = Trim(MSFhiv.TextMatrix(i, 4)) 'DTMNO
''''''                INSDTMTABQR(11) = time1 'TIME
''''''                INSDTMTABQR(12) = 0 'KEY
''''''                INSDTMTABQR(13) = 0 'KEY
''''''                INSDTMTABQR(14) = CInt(a) 'KEY
''''''                INSDTMTABQR.Execute
''''''
''''''
''''''                INSDMBCCOMMQR(0) = Trim(req_chglen(MSFhiv.TextMatrix(i, 1)))
''''''                If Len(Trim(txt_comments.Text)) > 0 Then
''''''                    INSDMBCCOMMQR(1) = Null
''''''                    INSDMBCCOMMQR(1) = txt_comments.Text
''''''                Else
''''''                    INSDMBCCOMMQR(1) = Null
''''''                End If
''''''                INSDMBCCOMMQR.Execute
''''''
''''''                UPDREQQR(0) = Trim(cmb_enterby.Text)
''''''                UPDREQQR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''''''                UPDREQQR(2) = "T"
''''''                UPDREQQR(3) = Trim(req_chglen(MSFhiv.TextMatrix(i, 1)))
''''''                UPDREQQR.Execute
''''''            End If
''''''        Next i
   End If
'commented as all three HBSAG(ELISA) ,HCV,HIV are shifted to Seriology Sundeep
    rdb.CommitTrans
    
    If ctr = 0 Then
        MsgBox "None of the data saved!!!", vbInformation, "DTM"
    Else
        MsgBox ctr & " record(s) saved successfully !!!", vbInformation, "DTM"
    End If
End If      'save

If flagsave = 2 Then    'Modify
    If Trim(cmb_option.Text) = "RCS" Then
        i = 1
        For i = 1 To MSFrcs.Rows - 1
            If MSFrcs.TextMatrix(i, 1) <> "" Then
                UPDDTMQRABO(0) = "" & Trim(MSFrcs.TextMatrix(i, 3))
                UPDDTMQRABO(1) = "" & Trim(MSFrcs.TextMatrix(i, 4))
                UPDDTMQRABO(2) = "" & Trim(MSFrcs.TextMatrix(i, 5))
                UPDDTMQRABO(3) = "" & Trim(MSFrcs.TextMatrix(i, 6))
                UPDDTMQRABO(4) = "" & Trim(MSFrcs.TextMatrix(i, 7))
                UPDDTMQRABO(5) = Trim(MSFrcs.TextMatrix(i, 2))
                UPDDTMQRABO(6) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
                'UPDDTMNO_QR(0) = "" & Trim(MSFrcs.TextMatrix(i, 3))
                'UPDDTMNO_QR(1) = "" & Trim(MSFrcs.TextMatrix(i, 4))
                UPDDTMNO_QR(0) = Trim(MSFrcs.TextMatrix(i, 3)) & " " & Trim(MSFrcs.TextMatrix(i, 4))
                UPDDTMNO_QR(1) = Trim(MSFrcs.TextMatrix(i, 2))
                UPDDTMNO_QR.Execute
               ' UPDDTMNO_QR(2) = "" & Trim(MSFrcs.TextMatrix(i, 5))
               'UPDDTMNO_QR(3) = "" & Trim(MSFrcs.TextMatrix(i, 6))
               'UPDDTMNO_QR(4) = "" & Trim(MSFrcs.TextMatrix(i, 7))
               'UPDDTMNO_QR(5) = Trim(MSFrcs.TextMatrix(i, 2))
                ' UPDDTMNO_QR(6) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
                UPDDTMQRABO.Execute
               ' UPDDTMNO_QR.Execute
                UPDREQQR(0) = Trim(cmb_enterby.Text)
                UPDREQQR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                UPDREQQR(2) = "T"
                UPDREQQR(3) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
                UPDREQQR.Execute
                'UPDDTMNO_QR(0) = Trim(Cmb_Enterby.Text)
                'UPDDTMNO_QR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                'UPDDTMNO_QR(2) = "T"
                'UPDDTMNO_QR(3) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
                UPDREQQR_P(0) = Trim(cmb_enterby.Text)
                UPDREQQR_P(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                UPDREQQR_P(2) = "T"
                UPDREQQR_P(3) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
                UPDREQQR_P.Execute
                
                'following txt_assigning is added by shweta 03-04-2013
            txt_comments.Text = Trim(MSFrcs.TextMatrix(i, 9))
            '''
                
                'UPDDTMNO_QR.Execute
                If Len(Trim(txt_comments.Text)) > 0 Then
                    UPDMBCCOMMQR(0) = Null
                    UPDMBCCOMMQR(0) = txt_comments.Text
                Else
                    UPDMBCCOMMQR(0) = Null
                End If
                UPDMBCCOMMQR(1) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
                UPDMBCCOMMQR.Execute
                txt_comments.Text = ""
            End If
        Next i
        'commented as all three HBSAG(ELISA) ,HCV,HIV are shifted to Seriology Sundeep
''''''    ElseIf Trim(cmb_option.Text) = "HBsAg(ELISA)" Then
''''''        If UPDDTMQRHBSAGFLAG = False Then
''''''            STR = "UPDATE DISLIB.DTMTAB SET HBSAG=?,DTMLAB=?,HBSAGKEY=? WHERE CASENO=? AND REQNO=?"
''''''            Set UPDDTMQRHBSAG = rdb.CreateQuery("UPDDTMHBSAGSTR", STR)
''''''            UPDDTMQRHBSAGFLAG = True
''''''        End If
''''''        i = 1
''''''        For i = 1 To MSFhbsAg.Rows - 1
''''''            If MSFhbsAg.TextMatrix(i, 1) <> "" Then
''''''                If Trim(MSFhbsAg.TextMatrix(i, 3)) = "" Then
''''''                    UPDDTMQRHBSAG(0) = Null
''''''                Else
''''''                    a = makekey
''''''                    UPDDTMQRHBSAG(0) = transform(Trim(MSFhbsAg.TextMatrix(i, 3)), a)
''''''                End If
''''''                UPDDTMQRHBSAG(1) = Trim(MSFhbsAg.TextMatrix(i, 4))
''''''                UPDDTMQRHBSAG(2) = CInt(a)
''''''                UPDDTMQRHBSAG(3) = Trim(MSFhbsAg.TextMatrix(i, 2))
''''''                UPDDTMQRHBSAG(4) = Trim(req_chglen(MSFhbsAg.TextMatrix(i, 1)))
''''''                UPDDTMQRHBSAG.Execute
''''''
''''''                UPDREQQR(0) = Trim(cmb_enterby.Text)
''''''                UPDREQQR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''''''                UPDREQQR(2) = "T"
''''''                UPDREQQR(3) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
''''''                UPDREQQR.Execute
''''''
''''''                If Len(Trim(txt_comments.Text)) > 0 Then
''''''                    UPDMBCCOMMQR(0) = Null
''''''                    UPDMBCCOMMQR(0) = txt_comments.Text
''''''                Else
''''''                    UPDMBCCOMMQR(0) = Null
''''''                End If
''''''                UPDMBCCOMMQR(1) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
''''''                UPDMBCCOMMQR.Execute
''''''            End If
''''''        Next i
''''''    ElseIf Trim(cmb_option.Text) = "HCV(ELISA)" Then
''''''        If UPDDTMQRHCVFLAG = False Then
''''''            STR = "UPDATE DISLIB.DTMTAB SET HCV=?,DTMLAB=?,HCVKEY=? WHERE CASENO=? AND REQNO=?"
''''''            Set UPDDTMQRHCV = rdb.CreateQuery("UPDDTMHCVSTR", STR)
''''''            UPDDTMQRHCVFLAG = True
''''''        End If
''''''        i = 1
''''''        For i = 1 To MSFhcv.Rows - 1
''''''            If MSFhcv.TextMatrix(i, 1) <> "" Then
''''''                If Trim(MSFhcv.TextMatrix(i, 3)) = "" Then
''''''                    UPDDTMQRHCV(0) = Null
''''''                Else
''''''                    a = makekey
''''''                    UPDDTMQRHCV(0) = transform(Trim(MSFhcv.TextMatrix(i, 3)), a)
''''''                End If
''''''                UPDDTMQRHCV(1) = "" & Trim(MSFhcv.TextMatrix(i, 4))
''''''                UPDDTMQRHCV(2) = CInt(a)
''''''                UPDDTMQRHCV(3) = Trim(MSFhcv.TextMatrix(i, 2))
''''''                UPDDTMQRHCV(4) = Trim(req_chglen(MSFhcv.TextMatrix(i, 1)))
''''''                UPDDTMQRHCV.Execute
''''''
''''''                UPDREQQR(0) = Trim(cmb_enterby.Text)
''''''                UPDREQQR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''''''                UPDREQQR(2) = "T"
''''''                UPDREQQR(3) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
''''''                UPDREQQR.Execute
''''''
''''''                If Len(Trim(txt_comments.Text)) > 0 Then
''''''                    UPDMBCCOMMQR(0) = Null
''''''                    UPDMBCCOMMQR(0) = txt_comments.Text
''''''                Else
''''''                    UPDMBCCOMMQR(0) = Null
''''''                End If
''''''                UPDMBCCOMMQR(1) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
''''''                UPDMBCCOMMQR.Execute
''''''            End If
''''''        Next i
''''''    ElseIf Trim(cmb_option.Text) = "HIV(ELISA)" Then
''''''        If UPDDTMQRHIVFLAG = False Then
''''''            STR = "UPDATE DISLIB.DTMTAB SET HIV=?,DTMLAB=?,HIVKEY=? WHERE CASENO=? AND REQNO=?"
''''''            Set UPDDTMQRHIV = rdb.CreateQuery("UPDDTMHIVSTR", STR)
''''''            UPDDTMQRHIVFLAG = True
''''''        End If
''''''        i = 1
''''''        For i = 1 To MSFhiv.Rows - 1
''''''            If MSFhiv.TextMatrix(i, 1) <> "" Then
''''''                If Trim(MSFhiv.TextMatrix(i, 3)) = "" Then
''''''                    UPDDTMQRHIV(0) = Null
''''''                Else
''''''                    a = makekey
''''''                    UPDDTMQRHIV(0) = transform(Trim(MSFhiv.TextMatrix(i, 3)), a)
''''''                End If
''''''                UPDDTMQRHIV(1) = "" & Trim(MSFhiv.TextMatrix(i, 4))
''''''                UPDDTMQRHIV(2) = CInt(a)
''''''                UPDDTMQRHIV(3) = Trim(MSFhiv.TextMatrix(i, 2))
''''''                UPDDTMQRHIV(4) = Trim(req_chglen(MSFhiv.TextMatrix(i, 1)))
''''''                UPDDTMQRHIV.Execute
''''''
''''''                UPDREQQR(0) = Trim(cmb_enterby.Text)
''''''                UPDREQQR(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
''''''                UPDREQQR(2) = "T"
''''''                UPDREQQR(3) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
''''''                UPDREQQR.Execute
''''''
''''''                If Len(Trim(txt_comments.Text)) > 0 Then
''''''                    UPDMBCCOMMQR(0) = Null
''''''                    UPDMBCCOMMQR(0) = txt_comments.Text
''''''                Else
''''''                    UPDMBCCOMMQR(0) = Null
''''''                End If
''''''                UPDMBCCOMMQR(1) = Trim(req_chglen(MSFrcs.TextMatrix(i, 1)))
''''''                UPDMBCCOMMQR.Execute
''''''            End If
''''''        Next i
'commented as all three HBSAG(ELISA) ,HCV,HIV are shifted to Seriology Sundeep
     End If
       
rdb.CommitTrans
MsgBox "Data is saved successfully !!!", vbInformation, "DTM"
End If
MODIFYFLAG = False
Exit Sub
'MSG:
rdb.RollbackTrans
'MsgBox Err.Description, vbCritical, "DTM"
Exit Sub
End Sub
Public Sub NEWALL()
Dim proc As String
Dim i
i = 1
If DrequisitnFLAG = False Then
    DrequisitnQR(0) = msk_fromdt.FormattedText
    DrequisitnQR(1) = msk_todt.FormattedText
    Set DrequisitnRS = DrequisitnQR.OpenResultset(1, 3)
    DrequisitnFLAG = True
Else
    DrequisitnQR(0) = msk_fromdt.FormattedText
    DrequisitnQR(1) = msk_todt.FormattedText
    'DrequisitnRS.MoreResults
    'DrequisitnRS.Requery
    Set DrequisitnRS = DrequisitnQR.OpenResultset(1, 3)
    DrequisitnFLAG = True
End If
If cmb_option.Text = "RCS" Then
    SSTab1.TabVisible(0) = True
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    MSFrcs.Rows = i
    If Not DrequisitnRS.EOF Then
        Do While Not DrequisitnRS.EOF
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = DrequisitnRS("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = DrequisitnRS("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = DrequisitnRS("reqno")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = DrequisitnRS("reqno")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H001" Or SELDMLPROCE_RS("PRO_CODE") = "H007" Or SELDMLPROCE_RS("PRO_CODE") = "H006" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                If SELDTMNO_RS.EOF Then
                    MSFrcs.AddItem i & vbTab & req_chglen(DrequisitnRS("reqno")) & vbTab & DrequisitnRS("CASENO")
                    i = i + 1
                Else
                    MSFrcs.AddItem i & vbTab & req_chglen(DrequisitnRS("reqno")) & vbTab & DrequisitnRS("CASENO") & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " " & Trim(SELDTMNO_RS("DTMNO"))
                    i = i + 1
                End If
            End If
            DrequisitnRS.MoveNext
        Loop
        End If
    ' commented as HBSAG(ELISA) is Shifted to  Sereiology Sundeep
'''''ElseIf cmb_option.Text = "HBsAg(ELISA)" Then
'''''    SSTab1.TabVisible(0) = False
'''''    SSTab1.TabVisible(1) = True
'''''    SSTab1.TabVisible(2) = False
'''''    SSTab1.TabVisible(3) = False
'''''    MSFhbsAg.Rows = i
'''''    If Not DrequisitnRS.EOF Then
'''''        Do While Not DrequisitnRS.EOF
'''''            If SELDTMNO_FLAG = False Then
'''''                SELDTMNO_QR(0) = DrequisitnRS("CASENO")
'''''                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
'''''                SELDTMNO_FLAG = True
'''''            Else
'''''                SELDTMNO_QR(0) = DrequisitnRS("CASENO")
'''''                SELDTMNO_RS.MoreResults
'''''                SELDTMNO_RS.Requery
'''''            End If
'''''            If SELDMLPROCE_FLAG = False Then
'''''                SELDMLPROCE_QR(0) = DrequisitnRS("reqno")
'''''                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
'''''                SELDMLPROCE_FLAG = True
'''''            Else
'''''                SELDMLPROCE_QR(0) = DrequisitnRS("reqno")
'''''                SELDMLPROCE_RS.MoreResults
'''''                SELDMLPROCE_RS.Requery
'''''            End If
'''''            proc = ""
'''''            If Not SELDMLPROCE_RS.EOF Then
'''''                Do While Not SELDMLPROCE_RS.EOF
'''''                    If SELDMLPROCE_RS("PRO_CODE") = "H012" Then
'''''                        proc = "T"
'''''                    End If
'''''                    SELDMLPROCE_RS.MoveNext
'''''                Loop
'''''            End If
'''''            If proc = "T" Then
'''''                If SELDTMNO_RS.EOF Then
'''''                    MSFhbsAg.AddItem i & vbTab & req_chglen(DrequisitnRS("REQNO")) & vbTab & DrequisitnRS("CASENO")
'''''                    i = i + 1
'''''                Else
'''''                    MSFhbsAg.AddItem i & vbTab & req_chglen(DrequisitnRS("REQNO")) & vbTab & DrequisitnRS("CASENO") & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
'''''                    i = i + 1
'''''                End If
'''''            End If
'''''            DrequisitnRS.MoveNext
'''''        Loop
'''''    End If
'''''ElseIf cmb_option.Text = "HCV(ELISA)" Then
'''''    SSTab1.TabVisible(0) = False
'''''    SSTab1.TabVisible(1) = False
'''''    SSTab1.TabVisible(2) = True
'''''    SSTab1.TabVisible(3) = False
'''''    MSFhcv.Rows = i
'''''    If Not DrequisitnRS.EOF Then
'''''        Do While Not DrequisitnRS.EOF
'''''            If SELDTMNO_FLAG = False Then
'''''                SELDTMNO_QR(0) = DrequisitnRS("CASENO")
'''''                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
'''''                SELDTMNO_FLAG = True
'''''            Else
'''''                SELDTMNO_QR(0) = DrequisitnRS("CASENO")
'''''                SELDTMNO_RS.MoreResults
'''''                SELDTMNO_RS.Requery
'''''            End If
'''''            If SELDMLPROCE_FLAG = False Then
'''''                SELDMLPROCE_QR(0) = DrequisitnRS("reqno")
'''''                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
'''''                SELDMLPROCE_FLAG = True
'''''            Else
'''''                SELDMLPROCE_QR(0) = DrequisitnRS("reqno")
'''''                SELDMLPROCE_RS.MoreResults
'''''                SELDMLPROCE_RS.Requery
'''''            End If
'''''            proc = ""
'''''            If Not SELDMLPROCE_RS.EOF Then
'''''                Do While Not SELDMLPROCE_RS.EOF
'''''                    If SELDMLPROCE_RS("PRO_CODE") = "H014" Then
'''''                        proc = "T"
'''''                    End If
'''''                    SELDMLPROCE_RS.MoveNext
'''''                Loop
'''''            End If
'''''            If proc = "T" Then
'''''                If SELDTMNO_RS.EOF Then
'''''                    MSFhcv.AddItem i & vbTab & req_chglen(DrequisitnRS("REQNO")) & vbTab & DrequisitnRS("CASENO")
'''''                    i = i + 1
'''''                Else
'''''                    MSFhcv.AddItem i & vbTab & req_chglen(DrequisitnRS("REQNO")) & vbTab & DrequisitnRS("CASENO") & vbTab & vbTab & "" & SELDTMNO_RS("DTMNO")
'''''                    i = i + 1
'''''                End If
'''''            End If
'''''            DrequisitnRS.MoveNext
'''''        Loop
'''''    End If
'''''ElseIf cmb_option.Text = "HIV(ELISA)" Then
'''''    SSTab1.TabVisible(0) = False
'''''    SSTab1.TabVisible(1) = False
'''''    SSTab1.TabVisible(2) = False
'''''    SSTab1.TabVisible(3) = True
'''''    MSFhiv.Rows = i
'''''    If Not DrequisitnRS.EOF Then
'''''        Do While Not DrequisitnRS.EOF
'''''            If SELDTMNO_FLAG = False Then
'''''                SELDTMNO_QR(0) = DrequisitnRS("CASENO")
'''''                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
'''''                SELDTMNO_FLAG = True
'''''            Else
'''''                SELDTMNO_QR(0) = DrequisitnRS("CASENO")
'''''                SELDTMNO_RS.MoreResults
'''''                SELDTMNO_RS.Requery
'''''            End If
'''''            If SELDMLPROCE_FLAG = False Then
'''''                SELDMLPROCE_QR(0) = DrequisitnRS("reqno")
'''''                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
'''''                SELDMLPROCE_FLAG = True
'''''            Else
'''''                SELDMLPROCE_QR(0) = DrequisitnRS("reqno")
'''''                SELDMLPROCE_RS.MoreResults
'''''                SELDMLPROCE_RS.Requery
'''''            End If
'''''            proc = ""
'''''            If Not SELDMLPROCE_RS.EOF Then
'''''                Do While Not SELDMLPROCE_RS.EOF
'''''                    If SELDMLPROCE_RS("PRO_CODE") = "H013" Then
'''''                        proc = "T"
'''''                    End If
'''''                    SELDMLPROCE_RS.MoveNext
'''''                Loop
'''''            End If
'''''            If proc = "T" Then
'''''                If SELDTMNO_RS.EOF Then
'''''                    MSFhiv.AddItem i & vbTab & req_chglen(DrequisitnRS("REQNO")) & vbTab & DrequisitnRS("CASENO")
'''''                    i = i + 1
'''''                Else
'''''                    MSFhiv.AddItem i & vbTab & req_chglen(DrequisitnRS("REQNO")) & vbTab & DrequisitnRS("CASENO") & vbTab & vbTab & "" & SELDTMNO_RS("DTMNO")
'''''                    i = i + 1
'''''                End If
'''''            End If
'''''            DrequisitnRS.MoveNext
'''''        Loop
'''''    End If
' commented as HBSAG(ELISA),HIV(ELISA) is Shifted to  Seriology Sundeep
End If
End Sub
Public Sub NEWCASE()
Dim proc As String
Dim i
i = 1
If selreflag = False Then
    selreQR(0) = Trim(req_chglen(Com_REQNO.Text))
    Set selreRS = selreQR.OpenResultset(1, 3)
    selreflag = True
Else
    selreQR(0) = Trim(req_chglen(Com_REQNO.Text))
    selreRS.MoreResults
    selreRS.Requery
End If
If SELDTMNO_FLAG = False Then
    SELDTMNO_QR(0) = msk_caseno.FormattedText
    Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
    SELDTMNO_FLAG = True
Else
    SELDTMNO_QR(0) = msk_caseno.FormattedText
    SELDTMNO_RS.MoreResults
    SELDTMNO_RS.Requery
End If
If cmb_option.Text = "RCS" Then
    SSTab1.TabVisible(0) = True
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    MSFrcs.Rows = i
    If Not selreRS.EOF Then
        Do While Not selreRS.EOF
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = selreRS("REQNO")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = selreRS("REQNO")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H001" Or SELDMLPROCE_RS("PRO_CODE") = "H007" Or SELDMLPROCE_RS("PRO_CODE") = "H006" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                If SELDTMNO_RS.EOF Then
                    MSFrcs.AddItem i & vbTab & req_chglen(selreRS("REQNO")) & vbTab & selreRS("CASENO")
                    i = i + 1
                Else
                    MSFrcs.AddItem i & vbTab & req_chglen(selreRS("REQNO")) & vbTab & selreRS("CASENO") & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "" & SELDTMNO_RS("DTMNO")
                    i = i + 1
                End If
            End If
            selreRS.MoveNext
        Loop
    End If
    'commented as all three HBSAG(ELISA) ,HCV,HIV are shifted to Seriology Sundeep
'''''ElseIf cmb_option.Text = "HBsAg(ELISA)" Then
'''''    SSTab1.TabVisible(0) = False
'''''    SSTab1.TabVisible(1) = True
'''''    SSTab1.TabVisible(2) = False
'''''    SSTab1.TabVisible(3) = False
'''''    MSFhbsAg.Rows = i
'''''    If Not selreRS.EOF Then
'''''        Do While Not selreRS.EOF
'''''            If SELDMLPROCE_FLAG = False Then
'''''                SELDMLPROCE_QR(0) = selreRS("REQNO")
'''''                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
'''''                SELDMLPROCE_FLAG = True
'''''            Else
'''''                SELDMLPROCE_QR(0) = selreRS("REQNO")
'''''                SELDMLPROCE_RS.MoreResults
'''''                SELDMLPROCE_RS.Requery
'''''            End If
'''''            proc = ""
'''''            If Not SELDMLPROCE_RS.EOF Then
'''''                Do While Not SELDMLPROCE_RS.EOF
'''''                    If SELDMLPROCE_RS("PRO_CODE") = "H012" Then
'''''                        proc = "T"
'''''                    End If
'''''                    SELDMLPROCE_RS.MoveNext
'''''                Loop
'''''            End If
'''''            If proc = "T" Then
'''''                If SELDTMNO_RS.EOF Then
'''''                    MSFhbsAg.AddItem i & vbTab & req_chglen(selreRS("REQNO")) & vbTab & selreRS("CASENO")
'''''                    i = i + 1
'''''                Else
'''''                    MSFhbsAg.AddItem i & vbTab & req_chglen(selreRS("REQNO")) & vbTab & selreRS("CASENO") & vbTab & vbTab & "" & SELDTMNO_RS("DTMNO")
'''''                    i = i + 1
'''''                End If
'''''            End If
'''''            selreRS.MoveNext
'''''        Loop
'''''    End If
'''''ElseIf cmb_option.Text = "HCV(ELISA)" Then
'''''    SSTab1.TabVisible(0) = False
'''''    SSTab1.TabVisible(1) = False
'''''    SSTab1.TabVisible(2) = True
'''''    SSTab1.TabVisible(3) = False
'''''    MSFhcv.Rows = i
'''''    If Not selreRS.EOF Then
'''''        Do While Not selreRS.EOF
'''''            If SELDMLPROCE_FLAG = False Then
'''''                SELDMLPROCE_QR(0) = selreRS("REQNO")
'''''                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
'''''                SELDMLPROCE_FLAG = True
'''''            Else
'''''                SELDMLPROCE_QR(0) = selreRS("REQNO")
'''''                SELDMLPROCE_RS.MoreResults
'''''                SELDMLPROCE_RS.Requery
'''''            End If
'''''            proc = ""
'''''            If Not SELDMLPROCE_RS.EOF Then
'''''                Do While Not SELDMLPROCE_RS.EOF
'''''                    If SELDMLPROCE_RS("PRO_CODE") = "H014" Then
'''''                        proc = "T"
'''''                    End If
'''''                    SELDMLPROCE_RS.MoveNext
'''''                Loop
'''''            End If
'''''            If proc = "T" Then
'''''                If SELDTMNO_RS.EOF Then
'''''                    MSFhcv.AddItem i & vbTab & req_chglen(selreRS("REQNO")) & vbTab & selreRS("CASENO")
'''''                    i = i + 1
'''''                Else
'''''                    MSFhcv.AddItem i & vbTab & req_chglen(selreRS("REQNO")) & vbTab & selreRS("CASENO") & vbTab & vbTab & "" & SELDTMNO_RS("DTMNO")
'''''                    i = i + 1
'''''                End If
'''''            End If
'''''            selreRS.MoveNext
'''''        Loop
'''''    End If
'''''ElseIf cmb_option.Text = "HIV(ELISA)" Then
'''''    SSTab1.TabVisible(0) = False
'''''    SSTab1.TabVisible(1) = False
'''''    SSTab1.TabVisible(2) = False
'''''    SSTab1.TabVisible(3) = True
'''''    MSFhiv.Rows = i
'''''    If Not selreRS.EOF Then
'''''        Do While Not selreRS.EOF
'''''            If SELDMLPROCE_FLAG = False Then
'''''                SELDMLPROCE_QR(0) = selreRS("REQNO")
'''''                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
'''''                SELDMLPROCE_FLAG = True
'''''            Else
'''''                SELDMLPROCE_QR(0) = selreRS("REQNO")
'''''                SELDMLPROCE_RS.MoreResults
'''''                SELDMLPROCE_RS.Requery
'''''            End If
'''''            proc = ""
'''''            If Not SELDMLPROCE_RS.EOF Then
'''''                Do While Not SELDMLPROCE_RS.EOF
'''''                    If SELDMLPROCE_RS("PRO_CODE") = "H013" Then
'''''                        proc = "T"
'''''                    End If
'''''                    SELDMLPROCE_RS.MoveNext
'''''                Loop
'''''            End If
'''''            If proc = "T" Then
'''''                If SELDTMNO_RS.EOF Then
'''''                    MSFhiv.AddItem i & vbTab & req_chglen(selreRS("REQNO")) & vbTab & selreRS("CASENO")
'''''                    i = i + 1
'''''                Else
'''''                    MSFhiv.AddItem i & vbTab & req_chglen(selreRS("REQNO")) & vbTab & selreRS("CASENO") & vbTab & vbTab & "" & SELDTMNO_RS("DTMNO")
'''''                    i = i + 1
'''''                End If
'''''            End If
'''''            selreRS.MoveNext
'''''        Loop
'''''    End If
'commented as all three HBSAG(ELISA) ,HCV,HIV are shifted to Seriology Sundeep
End If
End Sub
Public Sub VIEWCASE()
Dim proc As String
Dim i
i = 1
If finddtmFLAG = False Then
    finddtmQR(0) = Trim(req_chglen(Com_REQNO.Text))
    finddtmQR(1) = msk_caseno.FormattedText
    Set finddtmRS = finddtmQR.OpenResultset(1, 3)
    finddtmFLAG = True
Else
    finddtmQR(0) = Trim(req_chglen(Com_REQNO.Text))
    finddtmQR(1) = msk_caseno.FormattedText
    finddtmRS.MoreResults
    finddtmRS.Requery
End If
If Not finddtmRS.EOF Then
    If SELDTMNO_FLAG = False Then
        SELDTMNO_QR(0) = msk_caseno.FormattedText
        Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
        SELDTMNO_FLAG = True
    Else
        SELDTMNO_QR(0) = msk_caseno.FormattedText
        SELDTMNO_RS.MoreResults
        SELDTMNO_RS.Requery
    End If
    If selDTMTabFLAG = False Then
       selDTMTabQR(0) = Trim(req_chglen(Com_REQNO.Text))
       selDTMTabQR(1) = msk_caseno.FormattedText
       Set selDTMTabRS = selDTMTabQR.OpenResultset(1, 3)
       selDTMTabFLAG = True
    Else
        selDTMTabQR(0) = Trim(req_chglen(Com_REQNO.Text))
        selDTMTabQR(1) = msk_caseno.FormattedText
        selDTMTabRS.MoreResults
        selDTMTabRS.Requery
    End If
    If VARREQFLAG = False Then
        VARREQQR(0) = msk_caseno.FormattedText
        VARREQQR(1) = req_chglen(Com_REQNO.Text)
        Set VARREQRS = VARREQQR.OpenResultset(1, 3)
        VARREQFLAG = True
    Else
        VARREQQR(0) = msk_caseno.FormattedText
        VARREQQR(1) = req_chglen(Com_REQNO.Text)
        VARREQRS.MoreResults
        VARREQRS.Requery
    End If
    If Not VARREQRS.EOF Then
        cmb_enterby.Text = "" & Trim(VARREQRS("freportby"))
        'cmb_finalby.Text = "" & VARREQRS("PREPORTBY")
    End If
    If cmb_option.Text = "RCS" Then
        SSTab1.TabVisible(0) = True
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False
        SSTab1.TabVisible(3) = False
        MSFrcs.Rows = i
        If Not selDTMTabRS.EOF Then
            Do While Not selDTMTabRS.EOF
                If SELDMLPROCE_FLAG = False Then
                    SELDMLPROCE_QR(0) = selDTMTabRS("REQNO")
                    Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                    SELDMLPROCE_FLAG = True
                Else
                    SELDMLPROCE_QR(0) = selDTMTabRS("REQNO")
                    SELDMLPROCE_RS.MoreResults
                    SELDMLPROCE_RS.Requery
                End If
                proc = ""
                If Not SELDMLPROCE_RS.EOF Then
                    Do While Not SELDMLPROCE_RS.EOF
                        If SELDMLPROCE_RS("PRO_CODE") = "H001" Or SELDMLPROCE_RS("PRO_CODE") = "H007" Or SELDMLPROCE_RS("PRO_CODE") = "H006" Then
                            proc = "T"
                        End If
                        SELDMLPROCE_RS.MoveNext
                    Loop
                End If
                Dim COMEN  As String
                COMEN = ""
                If SELDMBC_FLAG = False Then
        SELDMBC_QR(0) = req_chglen(Com_REQNO.Text)
        Set SELDMBC_RS = SELDMBC_QR.OpenResultset(1, 3)
        SELDMBC_FLAG = True
    Else
        SELDMBC_QR(0) = req_chglen(Com_REQNO.Text)
        SELDMBC_RS.MoreResults
        SELDMBC_RS.Requery
    End If
    If Not SELDMBC_RS.EOF Then
        COMEN = "" & SELDMBC_RS("COMMENTS")
    End If
                
                If proc = "T" Then
                    If SELDTMNO_RS.EOF Then
                        MSFrcs.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS("REQNO"))) & vbTab & Trim(selDTMTabRS("CASENO")) & vbTab & "" & Trim(selDTMTabRS("abo")) & vbTab & "" & Trim(selDTMTabRS("rh")) & vbTab & "" & Trim(selDTMTabRS("abs")) & vbTab & "" & Trim(selDTMTabRS("dat")) & vbTab & "" & Trim(selDTMTabRS("iat")) & vbTab & "" & vbTab & COMEN '& vbTab & "View"
                        i = i + 1
                    Else
                        MSFrcs.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS("REQNO"))) & vbTab & Trim(selDTMTabRS("CASENO")) & vbTab & "" & Trim(selDTMTabRS("abo")) & vbTab & "" & Trim(selDTMTabRS("rh")) & vbTab & "" & Trim(selDTMTabRS("abs")) & vbTab & "" & Trim(selDTMTabRS("dat")) & vbTab & "" & Trim(selDTMTabRS("iat")) & vbTab & "" & SELDTMNO_RS("DTMNO") & vbTab & COMEN ''& vbTab & "View" '& vbTab & "View"
                        i = i + 1
                    End If
                End If
                selDTMTabRS.MoveNext
            Loop
        End If
        ''commented as all three HBSAG(ELISA) ,HCV,HIV are shifted to Seriology Sundeep
''''''''    ElseIf cmb_option.Text = "HBsAg(ELISA)" Then
''''''''        SSTab1.TabVisible(0) = False
''''''''        SSTab1.TabVisible(1) = True
''''''''        SSTab1.TabVisible(2) = False
''''''''        SSTab1.TabVisible(3) = False
''''''''        MSFhbsAg.Rows = i
''''''''        If Not selDTMTabRS.EOF Then
''''''''            Do While Not selDTMTabRS.EOF
''''''''                If SELDMLPROCE_FLAG = False Then
''''''''                    SELDMLPROCE_QR(0) = selDTMTabRS("REQNO")
''''''''                    Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
''''''''                    SELDMLPROCE_FLAG = True
''''''''                Else
''''''''                    SELDMLPROCE_QR(0) = selDTMTabRS("REQNO")
''''''''                    SELDMLPROCE_RS.MoreResults
''''''''                    SELDMLPROCE_RS.Requery
''''''''                End If
''''''''                proc = ""
''''''''                If Not SELDMLPROCE_RS.EOF Then
''''''''                    Do While Not SELDMLPROCE_RS.EOF
''''''''                        If SELDMLPROCE_RS("PRO_CODE") = "H012" Then
''''''''                            proc = "T"
''''''''                        End If
''''''''                        SELDMLPROCE_RS.MoveNext
''''''''                    Loop
''''''''                End If
''''''''                If proc = "T" Then
''''''''                    a = CInt(selDTMTabRS("hbsagkey"))
''''''''                    If SELDTMNO_RS.EOF Then
''''''''                        If IsNull(selDTMTabRS("hbsag")) Then
''''''''                            MSFhbsAg.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS("REQNO"))) & vbTab & Trim(selDTMTabRS("CASENO"))
''''''''                        Else
''''''''                            MSFhbsAg.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS("REQNO"))) & vbTab & Trim(selDTMTabRS("CASENO")) & vbTab & DECRPT(Trim(selDTMTabRS("hbsag")), a)
''''''''                        End If
''''''''                        i = i + 1
''''''''                    Else
''''''''                        If IsNull(selDTMTabRS("hbsag")) Then
''''''''                            MSFhbsAg.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS("REQNO"))) & vbTab & Trim(selDTMTabRS("CASENO")) & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
''''''''                        Else
''''''''                            MSFhbsAg.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS("REQNO"))) & vbTab & Trim(selDTMTabRS("CASENO")) & vbTab & DECRPT(Trim(selDTMTabRS("hbsag")), a) & vbTab & "" & SELDTMNO_RS("DTMNO")
''''''''                        End If
''''''''                        i = i + 1
''''''''                    End If
''''''''                End If
''''''''                selDTMTabRS.MoveNext
''''''''            Loop
''''''''        End If
''''''''    ElseIf cmb_option.Text = "HCV(ELISA)" Then
''''''''        SSTab1.TabVisible(0) = False
''''''''        SSTab1.TabVisible(1) = False
''''''''        SSTab1.TabVisible(2) = True
''''''''        SSTab1.TabVisible(3) = False
''''''''        MSFhcv.Rows = i
''''''''        If Not selDTMTabRS.EOF Then
''''''''            Do While Not selDTMTabRS.EOF
''''''''                If SELDMLPROCE_FLAG = False Then
''''''''                    SELDMLPROCE_QR(0) = selDTMTabRS("REQNO")
''''''''                    Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
''''''''                    SELDMLPROCE_FLAG = True
''''''''                Else
''''''''                    SELDMLPROCE_QR(0) = selDTMTabRS("REQNO")
''''''''                    SELDMLPROCE_RS.MoreResults
''''''''                    SELDMLPROCE_RS.Requery
''''''''                End If
''''''''                proc = ""
''''''''                If Not SELDMLPROCE_RS.EOF Then
''''''''                    Do While Not SELDMLPROCE_RS.EOF
''''''''                        If SELDMLPROCE_RS("PRO_CODE") = "H014" Then
''''''''                            proc = "T"
''''''''                        End If
''''''''                        SELDMLPROCE_RS.MoveNext
''''''''                    Loop
''''''''                End If
''''''''                If proc = "T" Then
''''''''                    a = CInt(selDTMTabRS("HCVkey"))
''''''''                    If SELDTMNO_RS.EOF Then
''''''''                        If IsNull(selDTMTabRS("hcv")) Then
''''''''                            MSFhcv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS("REQNO"))) & vbTab & Trim(selDTMTabRS("CASENO"))
''''''''                        Else
''''''''                            MSFhcv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS("REQNO"))) & vbTab & Trim(selDTMTabRS("CASENO")) & vbTab & DECRPT(Trim(selDTMTabRS("hcv")), a)
''''''''                        End If
''''''''                        i = i + 1
''''''''                    Else
''''''''                        If IsNull(selDTMTabRS("hcv")) Then
''''''''                            MSFhcv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS("REQNO"))) & vbTab & Trim(selDTMTabRS("CASENO")) & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
''''''''                        Else
''''''''                            MSFhcv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS("REQNO"))) & vbTab & Trim(selDTMTabRS("CASENO")) & vbTab & DECRPT(Trim(selDTMTabRS("hcv")), a) & vbTab & "" & SELDTMNO_RS("DTMNO")
''''''''                        End If
''''''''                        i = i + 1
''''''''                    End If
''''''''                End If
''''''''                selDTMTabRS.MoveNext
''''''''            Loop
''''''''        End If
''''''''    ElseIf cmb_option.Text = "HIV(ELISA)" Then
''''''''        SSTab1.TabVisible(0) = False
''''''''        SSTab1.TabVisible(1) = False
''''''''        SSTab1.TabVisible(2) = False
''''''''        SSTab1.TabVisible(3) = True
''''''''        MSFhiv.Rows = i
''''''''        If Not selDTMTabRS.EOF Then
''''''''            Do While Not selDTMTabRS.EOF
''''''''                If SELDMLPROCE_FLAG = False Then
''''''''                    SELDMLPROCE_QR(0) = selDTMTabRS("REQNO")
''''''''                    Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
''''''''                    SELDMLPROCE_FLAG = True
''''''''                Else
''''''''                    SELDMLPROCE_QR(0) = selDTMTabRS("REQNO")
''''''''                    SELDMLPROCE_RS.MoreResults
''''''''                    SELDMLPROCE_RS.Requery
''''''''                End If
''''''''                proc = ""
''''''''                If Not SELDMLPROCE_RS.EOF Then
''''''''                    Do While Not SELDMLPROCE_RS.EOF
''''''''                        If SELDMLPROCE_RS("PRO_CODE") = "H013" Then
''''''''                            proc = "T"
''''''''                        End If
''''''''                        SELDMLPROCE_RS.MoveNext
''''''''                    Loop
''''''''                End If
''''''''                If proc = "T" Then
''''''''                    a = CInt(selDTMTabRS("HIVkey"))
''''''''                    If SELDTMNO_RS.EOF Then
''''''''                        If IsNull(selDTMTabRS("hiv")) Then
''''''''                            MSFhiv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS("REQNO"))) & vbTab & Trim(selDTMTabRS("CASENO"))
''''''''                        Else
''''''''                            MSFhiv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS("REQNO"))) & vbTab & Trim(selDTMTabRS("CASENO")) & vbTab & DECRPT(Trim(selDTMTabRS("hiv")), a)
''''''''                        End If
''''''''                        i = i + 1
''''''''                    Else
''''''''                        MSFhiv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS("REQNO"))) & vbTab & Trim(selDTMTabRS("CASENO")) & vbTab & Trim(selDTMTabRS("hiv")) & vbTab & "" & Trim(SELDTMNO_RS("DTMNO"))
''''''''                        i = i + 1
''''''''                    End If
''''''''                End If
''''''''                selDTMTabRS.MoveNext
''''''''            Loop
''''''''        End If
'commented as all three HBSAG(ELISA) ,HCV,HIV are shifted to Seriology Sundeep
    End If ''*******
'End If
    If SELDMBC_FLAG = False Then
        SELDMBC_QR(0) = req_chglen(Com_REQNO.Text)
        Set SELDMBC_RS = SELDMBC_QR.OpenResultset(1, 3)
        SELDMBC_FLAG = True
    Else
        SELDMBC_QR(0) = req_chglen(Com_REQNO.Text)
        SELDMBC_RS.MoreResults
        SELDMBC_RS.Requery
    End If
    If Not SELDMBC_RS.EOF Then
        txt_comments.Text = "" & SELDMBC_RS("COMMENTS")
    End If
''***commit
    If REQ2_QRFLAG = False Then
        req2_qr(0) = req_chglen(Trim(Com_REQNO.Text))
        Set req2_rs = req2_qr.OpenResultset(1, 3)
        REQ2_QRFLAG = True
    Else
        req2_qr(0) = req_chglen(Trim(Com_REQNO.Text))
        req2_rs.MoreResults
        req2_rs.Requery
    End If
    If Not req2_rs.EOF Then
        If req2_rs("COMMITFLAG") = "Y" Then
            MsgBox " Report is Committed ....U can't modify it ! "
            btnenabl ("SAVE")
            cmd_commit.Enabled = False
            Exit Sub
        End If
    End If
End If
End Sub
Public Sub AWAITEDALL()
Dim proc As String
Dim i
i = 1
If cmb_option.Text = "RCS" Then
    SSTab1.TabVisible(0) = True
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    If Drequisitn_FLAG_RCS = False Then
       Drequisitn_QR_RCS(0) = msk_fromdt.FormattedText
       Drequisitn_QR_RCS(1) = msk_todt.FormattedText
       Set Drequisitn_RS_RCS = Drequisitn_QR_RCS.OpenResultset(1, 3)
       Drequisitn_FLAG_RCS = True
    Else
       Drequisitn_QR_RCS(0) = msk_fromdt.FormattedText
       Drequisitn_QR_RCS(1) = msk_todt.FormattedText
       Drequisitn_RS_RCS.MoreResults
       Drequisitn_RS_RCS.Requery
    End If
    MSFrcs.Rows = i
    If Not Drequisitn_RS_RCS.EOF Then
        Do While Not Drequisitn_RS_RCS.EOF
            If selDTMTabFLAG1 = False Then
               selDTMTabQR1(0) = Trim(Drequisitn_RS_RCS("REQNO"))
               Set selDTMTabRS1 = selDTMTabQR1.OpenResultset(1, 3)
               selDTMTabFLAG1 = True
            Else
                selDTMTabQR1(0) = Trim(Drequisitn_RS_RCS("REQNO"))
                selDTMTabRS1.MoreResults
                selDTMTabRS1.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS_RCS("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS_RCS("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = Drequisitn_RS_RCS("reqno")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = Drequisitn_RS_RCS("reqno")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H001" Or SELDMLPROCE_RS("PRO_CODE") = "H007" Or SELDMLPROCE_RS("PRO_CODE") = "H006" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                If SELDTMNO_RS.EOF Then
                    MSFrcs.AddItem i & vbTab & req_chglen(Drequisitn_RS_RCS("reqno")) & vbTab & Drequisitn_RS_RCS("CASENO") & vbTab & "" & Trim(selDTMTabRS1("abo")) & vbTab & "" & Trim(selDTMTabRS1("rh")) & vbTab & "" & Trim(selDTMTabRS1("abs")) & vbTab & "" & Trim(selDTMTabRS1("dat")) & vbTab & "" & Trim(selDTMTabRS1("iat"))
                    i = i + 1
                Else
                    MSFrcs.AddItem i & vbTab & req_chglen(Drequisitn_RS_RCS("reqno")) & vbTab & Drequisitn_RS_RCS("CASENO") & vbTab & "" & Trim(selDTMTabRS1("abo")) & vbTab & "" & Trim(selDTMTabRS1("rh")) & vbTab & "" & Trim(selDTMTabRS1("abs")) & vbTab & "" & Trim(selDTMTabRS1("dat")) & vbTab & "" & Trim(selDTMTabRS1("iat")) & vbTab & " " & Trim(SELDTMNO_RS("DTMNO"))
                    i = i + 1
                End If
            End If
            Drequisitn_RS_RCS.MoveNext
        Loop
    End If
ElseIf cmb_option.Text = "HBsAg(ELISA)" Then
    SSTab1.TabVisible(0) = False
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    If Drequisitn_FLAG_HBSAG = False Then
       Drequisitn_QR_HBSAG(0) = msk_fromdt.FormattedText
       Drequisitn_QR_HBSAG(1) = msk_todt.FormattedText
       Set Drequisitn_RS_HBSAG = Drequisitn_QR_HBSAG.OpenResultset(1, 3)
       Drequisitn_FLAG_HBSAG = True
    Else
       Drequisitn_QR_HBSAG(0) = msk_fromdt.FormattedText
       Drequisitn_QR_HBSAG(1) = msk_todt.FormattedText
       Drequisitn_RS_HBSAG.MoreResults
       Drequisitn_RS_HBSAG.Requery
    End If
    MSFhbsAg.Rows = i
    If Not Drequisitn_RS_HBSAG.EOF Then
        Do While Not Drequisitn_RS_HBSAG.EOF
            If selDTMTabFLAG1 = False Then
               selDTMTabQR1(0) = Trim(Drequisitn_RS_HBSAG("REQNO"))
               Set selDTMTabRS1 = selDTMTabQR1.OpenResultset(1, 3)
               selDTMTabFLAG1 = True
            Else
                selDTMTabQR1(0) = Trim(Drequisitn_RS_HBSAG("REQNO"))
                selDTMTabRS1.MoreResults
                selDTMTabRS1.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS_HBSAG("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS_HBSAG("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = Drequisitn_RS_HBSAG("reqno")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = Drequisitn_RS_HBSAG("reqno")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H012" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                If SELDTMNO_RS.EOF Then
                    MSFhbsAg.AddItem i & vbTab & req_chglen(Drequisitn_RS_HBSAG("REQNO")) & vbTab & Drequisitn_RS_HBSAG("CASENO") & vbTab & "" & selDTMTabRS1("hbsag")
                    i = i + 1
                Else
                    MSFhbsAg.AddItem i & vbTab & req_chglen(Drequisitn_RS_HBSAG("REQNO")) & vbTab & Drequisitn_RS_HBSAG("CASENO") & vbTab & "" & selDTMTabRS1("hbsag") & vbTab & "" & SELDTMNO_RS("DTMNO")
                    i = i + 1
                End If
            End If
            Drequisitn_RS_HBSAG.MoveNext
        Loop
    End If
ElseIf cmb_option.Text = "HCV(ELISA)" Then
    SSTab1.TabVisible(0) = False
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = True
    SSTab1.TabVisible(3) = False
    If Drequisitn_FLAG_HCV = False Then
       Drequisitn_QR_HCV(0) = msk_fromdt.FormattedText
       Drequisitn_QR_HCV(1) = msk_todt.FormattedText
       Set Drequisitn_RS_HCV = Drequisitn_QR_HCV.OpenResultset(1, 3)
       Drequisitn_FLAG_HCV = True
    Else
       Drequisitn_QR_HCV(0) = msk_fromdt.FormattedText
       Drequisitn_QR_HCV(1) = msk_todt.FormattedText
       Drequisitn_RS_HCV.MoreResults
       Drequisitn_RS_HCV.Requery
    End If
    MSFhcv.Rows = i
    If Not Drequisitn_RS_HCV.EOF Then
        Do While Not Drequisitn_RS_HCV.EOF
            If selDTMTabFLAG1 = False Then
               selDTMTabQR1(0) = Trim(Drequisitn_RS_HCV("REQNO"))
               Set selDTMTabRS1 = selDTMTabQR1.OpenResultset(1, 3)
               selDTMTabFLAG1 = True
            Else
                selDTMTabQR1(0) = Trim(Drequisitn_RS_HCV("REQNO"))
                selDTMTabRS1.MoreResults
                selDTMTabRS1.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS_HCV("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS_HCV("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = Drequisitn_RS_HCV("reqno")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = Drequisitn_RS_HCV("reqno")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H014" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                If SELDTMNO_RS.EOF Then
                    MSFhcv.AddItem i & vbTab & req_chglen(Drequisitn_RS_HCV("REQNO")) & vbTab & Drequisitn_RS_HCV("CASENO") & vbTab & "" & selDTMTabRS1("hcv")
                    i = i + 1
                Else
                    MSFhcv.AddItem i & vbTab & req_chglen(Drequisitn_RS_HCV("REQNO")) & vbTab & Drequisitn_RS_HCV("CASENO") & vbTab & "" & selDTMTabRS1("hcv") & vbTab & SELDTMNO_RS("DTMNO")
                    i = i + 1
                End If
            End If
            Drequisitn_RS_HCV.MoveNext
        Loop
    End If
ElseIf cmb_option.Text = "HIV(ELISA)" Then
    SSTab1.TabVisible(0) = False
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = True
    If Drequisitn_FLAG_HIV = False Then
       Drequisitn_QR_HIV(0) = msk_fromdt.FormattedText
       Drequisitn_QR_HIV(1) = msk_todt.FormattedText
       Set Drequisitn_RS_HIV = Drequisitn_QR_HIV.OpenResultset(1, 3)
       Drequisitn_FLAG_HIV = True
    Else
       Drequisitn_QR_HIV(0) = msk_fromdt.FormattedText
       Drequisitn_QR_HIV(1) = msk_todt.FormattedText
       Drequisitn_RS_HIV.MoreResults
       Drequisitn_RS_HIV.Requery
    End If
    MSFhiv.Rows = i
    If Not Drequisitn_RS_HIV.EOF Then
        Do While Not Drequisitn_RS_HIV.EOF
            If selDTMTabFLAG1 = False Then
               selDTMTabQR1(0) = Trim(Drequisitn_RS_HIV("REQNO"))
               Set selDTMTabRS1 = selDTMTabQR1.OpenResultset(1, 3)
               selDTMTabFLAG1 = True
            Else
                selDTMTabQR1(0) = Trim(Drequisitn_RS_HIV("REQNO"))
                selDTMTabRS1.MoreResults
                selDTMTabRS1.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS_HIV("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS_HIV("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = Drequisitn_RS_HIV("reqno")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = Drequisitn_RS_HIV("reqno")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H013" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                If SELDTMNO_RS.EOF Then
                    MSFhiv.AddItem i & vbTab & req_chglen(Drequisitn_RS_HIV("REQNO")) & vbTab & Drequisitn_RS_HIV("CASENO") & vbTab & "" & Trim(selDTMTabRS1("hiv"))
                    i = i + 1
                Else
                    MSFhiv.AddItem i & vbTab & req_chglen(Drequisitn_RS_HIV("REQNO")) & vbTab & Drequisitn_RS_HIV("CASENO") & vbTab & "" & Trim(selDTMTabRS1("hiv")) & vbTab & "" & SELDTMNO_RS("DTMNO")
                    i = i + 1
                End If
            End If
            Drequisitn_RS_HIV.MoveNext
        Loop
    End If
End If
End Sub
Public Sub VIEWALL()
Dim proc As String
Dim i
i = 1
If Drequisitn_FLAG = False Then
    Drequisitn_QR(0) = msk_fromdt.FormattedText
    Drequisitn_QR(1) = msk_todt.FormattedText
    Drequisitn_QR(2) = FTIME.FormattedText
    Drequisitn_QR(3) = TTIME.FormattedText
    Set Drequisitn_RS = Drequisitn_QR.OpenResultset(1, 3)
    Drequisitn_FLAG = True
Else
    Drequisitn_QR(0) = msk_fromdt.FormattedText
    Drequisitn_QR(1) = msk_todt.FormattedText
    Drequisitn_QR(2) = FTIME.FormattedText
    Drequisitn_QR(3) = TTIME.FormattedText
    Drequisitn_RS.MoreResults
    Drequisitn_RS.Requery
End If
If cmb_option.Text = "RCS" Then
    SSTab1.TabVisible(0) = True
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    MSFrcs.Rows = i
    If Not Drequisitn_RS.EOF Then
        Do While Not Drequisitn_RS.EOF
            If selDTMTabFLAG1 = False Then
               selDTMTabQR1(0) = Trim(Drequisitn_RS("REQNO"))
               Set selDTMTabRS1 = selDTMTabQR1.OpenResultset(1, 3)
               selDTMTabFLAG1 = True
            Else
                selDTMTabQR1(0) = Trim(Drequisitn_RS("REQNO"))
                selDTMTabRS1.MoreResults
                selDTMTabRS1.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = Drequisitn_RS("reqno")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = Drequisitn_RS("reqno")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H001" Or SELDMLPROCE_RS("PRO_CODE") = "H007" Or SELDMLPROCE_RS("PRO_CODE") = "H006" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                If SELDTMNO_RS.EOF Then
                    'If Trim(selDTMTabRS1("abo")) <> "" Or Trim(selDTMTabRS1("rh")) <> "" Or Trim(selDTMTabRS1("abs")) <> "" Or Trim(selDTMTabRS1("dat")) <> "" Or Trim(selDTMTabRS1("iat")) Then
                    If Not selDTMTabRS1.EOF Then
                        MSFrcs.AddItem i & vbTab & req_chglen(Drequisitn_RS("reqno")) & vbTab & Drequisitn_RS("CASENO") & vbTab & "" & Trim(selDTMTabRS1("abo")) & vbTab & "" & Trim(selDTMTabRS1("rh")) & vbTab & "" & Trim(selDTMTabRS1("abs")) & vbTab & "" & Trim(selDTMTabRS1("dat")) & vbTab & "" & Trim(selDTMTabRS1("iat")) '& vbTab & "VIEW"
                    End If
                    i = i + 1
                Else
                    'If Trim(selDTMTabRS1("abo")) <> "" Or Trim(selDTMTabRS1("rh")) <> "" Or Trim(selDTMTabRS1("abs")) <> "" Or Trim(selDTMTabRS1("dat")) <> "" Or Trim(selDTMTabRS1("iat")) Then
                    If Not selDTMTabRS1.EOF Then
                        MSFrcs.AddItem i & vbTab & req_chglen(Drequisitn_RS("reqno")) & vbTab & Drequisitn_RS("CASENO") & vbTab & "" & Trim(selDTMTabRS1("abo")) & vbTab & "" & Trim(selDTMTabRS1("rh")) & vbTab & "" & Trim(selDTMTabRS1("abs")) & vbTab & "" & Trim(selDTMTabRS1("dat")) & vbTab & "" & Trim(selDTMTabRS1("iat")) & vbTab & " " & Trim(SELDTMNO_RS("DTMNO")) '& vbTab & "VIEW"
                    End If
                    i = i + 1
                End If
            End If
            Drequisitn_RS.MoveNext
        Loop
    End If
ElseIf cmb_option.Text = "HBsAg(ELISA)" Then
    SSTab1.TabVisible(0) = False
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    MSFhbsAg.Rows = i
    If Not Drequisitn_RS.EOF Then
        Do While Not Drequisitn_RS.EOF
            If selDTMTabFLAG1 = False Then
               selDTMTabQR1(0) = Trim(Drequisitn_RS("REQNO"))
               Set selDTMTabRS1 = selDTMTabQR1.OpenResultset(1, 3)
               selDTMTabFLAG1 = True
            Else
                selDTMTabQR1(0) = Trim(Drequisitn_RS("REQNO"))
                selDTMTabRS1.MoreResults
                selDTMTabRS1.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = Drequisitn_RS("reqno")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = Drequisitn_RS("reqno")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H012" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                a = CInt(selDTMTabRS1("hbsagKEY"))
                If SELDTMNO_RS.EOF Then
                    If IsNull(selDTMTabRS1("hbsag")) Then
                        MSFhbsAg.AddItem i & vbTab & req_chglen(Drequisitn_RS("REQNO")) & vbTab & Drequisitn_RS("CASENO")
                    Else
                        MSFhbsAg.AddItem i & vbTab & req_chglen(Drequisitn_RS("REQNO")) & vbTab & Drequisitn_RS("CASENO") & vbTab & "" & DECRPT(Trim(selDTMTabRS1("hbsag")), a)
                    End If
                    i = i + 1
                Else
                    If IsNull(selDTMTabRS1("hbsag")) Then
                        MSFhbsAg.AddItem i & vbTab & req_chglen(Drequisitn_RS("REQNO")) & vbTab & Drequisitn_RS("CASENO") & vbTab & "" & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                    Else
                        MSFhbsAg.AddItem i & vbTab & req_chglen(Drequisitn_RS("REQNO")) & vbTab & Drequisitn_RS("CASENO") & vbTab & "" & DECRPT(Trim(selDTMTabRS1("hbsag")), a) & vbTab & "" & SELDTMNO_RS("DTMNO")
                    End If
                    i = i + 1
                End If
            End If
            Drequisitn_RS.MoveNext
        Loop
    End If
ElseIf cmb_option.Text = "HCV(ELISA)" Then
    SSTab1.TabVisible(0) = False
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = True
    SSTab1.TabVisible(3) = False
    MSFhcv.Rows = i
    If Not Drequisitn_RS.EOF Then
        Do While Not Drequisitn_RS.EOF
            If selDTMTabFLAG1 = False Then
               selDTMTabQR1(0) = Trim(Drequisitn_RS("REQNO"))
               Set selDTMTabRS1 = selDTMTabQR1.OpenResultset(1, 3)
               selDTMTabFLAG1 = True
            Else
                selDTMTabQR1(0) = Trim(Drequisitn_RS("REQNO"))
                selDTMTabRS1.MoreResults
                selDTMTabRS1.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = Drequisitn_RS("reqno")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = Drequisitn_RS("reqno")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H014" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                a = CInt(selDTMTabRS1("HCVKEY"))
                If SELDTMNO_RS.EOF Then
                    If IsNull(selDTMTabRS1("hcv")) Then
                        MSFhcv.AddItem i & vbTab & req_chglen(Drequisitn_RS("REQNO")) & vbTab & Drequisitn_RS("CASENO")
                    Else
                        MSFhcv.AddItem i & vbTab & req_chglen(Drequisitn_RS("REQNO")) & vbTab & Drequisitn_RS("CASENO") & vbTab & DECRPT(Trim(selDTMTabRS1("hcv")), a)
                    End If
                    i = i + 1
                Else
                    If IsNull(selDTMTabRS1("hcv")) Then
                        MSFhcv.AddItem i & vbTab & req_chglen(Drequisitn_RS("REQNO")) & vbTab & Drequisitn_RS("CASENO") & vbTab & "" & vbTab & SELDTMNO_RS("DTMNO")
                    Else
                        MSFhcv.AddItem i & vbTab & req_chglen(Drequisitn_RS("REQNO")) & vbTab & Drequisitn_RS("CASENO") & vbTab & "" & DECRPT(Trim(selDTMTabRS1("hcv")), a) & vbTab & SELDTMNO_RS("DTMNO")
                    End If
                    i = i + 1
                End If
            End If
            Drequisitn_RS.MoveNext
        Loop
    End If
ElseIf cmb_option.Text = "HIV(ELISA)" Then
    SSTab1.TabVisible(0) = False
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = True
    MSFhiv.Rows = i
    If Not Drequisitn_RS.EOF Then
        Do While Not Drequisitn_RS.EOF
            If selDTMTabFLAG1 = False Then
               selDTMTabQR1(0) = Trim(Drequisitn_RS("REQNO"))
               Set selDTMTabRS1 = selDTMTabQR1.OpenResultset(1, 3)
               selDTMTabFLAG1 = True
            Else
                selDTMTabQR1(0) = Trim(Drequisitn_RS("REQNO"))
                selDTMTabRS1.MoreResults
                selDTMTabRS1.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = Drequisitn_RS("reqno")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = Drequisitn_RS("reqno")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H013" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                a = CInt(selDTMTabRS1("hivkey"))
                If SELDTMNO_RS.EOF Then
                    If IsNull(selDTMTabRS1("hiv")) Then
                        MSFhiv.AddItem i & vbTab & req_chglen(Drequisitn_RS("REQNO")) & vbTab & Drequisitn_RS("CASENO")
                    Else
                        MSFhiv.AddItem i & vbTab & req_chglen(Drequisitn_RS("REQNO")) & vbTab & Drequisitn_RS("CASENO") & vbTab & "" & DECRPT(Trim(selDTMTabRS1("hiv")), a)
                    End If
                    i = i + 1
                Else
                    If IsNull(selDTMTabRS1("hiv")) Then
                        MSFhiv.AddItem i & vbTab & req_chglen(Drequisitn_RS("REQNO")) & vbTab & Drequisitn_RS("CASENO") & vbTab & "" & vbTab & "" & SELDTMNO_RS("DTMNO")
                    Else
                        MSFhiv.AddItem i & vbTab & req_chglen(Drequisitn_RS("REQNO")) & vbTab & Drequisitn_RS("CASENO") & vbTab & "" & DECRPT(Trim(selDTMTabRS1("hiv")), a) & vbTab & "" & SELDTMNO_RS("DTMNO")
                    End If
                    i = i + 1
                End If
            End If
            Drequisitn_RS.MoveNext
        Loop
    End If
End If
End Sub
Public Sub AWAITEDCASE()
Dim proc As String
Dim i
i = 1
If finddtmFLAG = False Then
    finddtmQR(0) = Trim(req_chglen(Com_REQNO.Text))
    finddtmQR(1) = msk_caseno.FormattedText
    Set finddtmRS = finddtmQR.OpenResultset(1, 3)
    finddtmFLAG = True
Else
    finddtmQR(0) = Trim(req_chglen(Com_REQNO.Text))
    finddtmQR(1) = msk_caseno.FormattedText
    finddtmRS.MoreResults
    finddtmRS.Requery
End If
If Not finddtmRS.EOF Then
    If SELDTMNO_FLAG = False Then
        SELDTMNO_QR(0) = msk_caseno.FormattedText
        Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
        SELDTMNO_FLAG = True
    Else
        SELDTMNO_QR(0) = msk_caseno.FormattedText
        SELDTMNO_RS.MoreResults
        SELDTMNO_RS.Requery
    End If
    If selDTMTabFLAG2 = False Then
       selDTMTabQR2(0) = Trim(req_chglen(Com_REQNO.Text))
       selDTMTabQR2(1) = msk_caseno.FormattedText
       Set selDTMTabRS2 = selDTMTabQR2.OpenResultset(1, 3)
       selDTMTabFLAG2 = True
    Else
        selDTMTabQR2(0) = Trim(req_chglen(Com_REQNO.Text))
        selDTMTabQR2(1) = msk_caseno.FormattedText
        selDTMTabRS2.MoreResults
        selDTMTabRS2.Requery
    End If
    If VARREQFLAG = False Then
        VARREQQR(0) = msk_caseno.FormattedText
        VARREQQR(1) = req_chglen(Com_REQNO.Text)
        Set VARREQRS = VARREQQR.OpenResultset(1, 3)
        VARREQFLAG = True
    Else
        VARREQQR(0) = msk_caseno.FormattedText
        VARREQQR(1) = req_chglen(Com_REQNO.Text)
        VARREQRS.MoreResults
        VARREQRS.Requery
    End If
    If Not VARREQRS.EOF Then
        cmb_enterby.Text = "" & VARREQRS("freportby")
        'cmb_finalby.Text = "" & VARREQRS("PREPORTBY")
    End If
    
    If cmb_option.Text = "RCS" Then
        SSTab1.TabVisible(0) = True
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False
        SSTab1.TabVisible(3) = False
        MSFrcs.Rows = i
        If Not selDTMTabRS2.EOF Then
            Do While Not selDTMTabRS2.EOF
                If SELDMLPROCE_FLAG = False Then
                    SELDMLPROCE_QR(0) = selDTMTabRS2("REQNO")
                    Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                    SELDMLPROCE_FLAG = True
                Else
                    SELDMLPROCE_QR(0) = selDTMTabRS2("REQNO")
                    SELDMLPROCE_RS.MoreResults
                    SELDMLPROCE_RS.Requery
                End If
                proc = ""
                If Not SELDMLPROCE_RS.EOF Then
                    Do While Not SELDMLPROCE_RS.EOF
                        If SELDMLPROCE_RS("PRO_CODE") = "H001" Or SELDMLPROCE_RS("PRO_CODE") = "H007" Or SELDMLPROCE_RS("PRO_CODE") = "H006" Then
                            proc = "T"
                        End If
                        SELDMLPROCE_RS.MoveNext
                    Loop
                End If
                If proc = "T" Then
                    If SELDTMNO_RS.EOF Then
                        MSFrcs.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO")) & vbTab & "" & Trim(selDTMTabRS2("abo")) & vbTab & "" & Trim(selDTMTabRS2("rh")) & vbTab & "" & Trim(selDTMTabRS2("abs")) & vbTab & "" & Trim(selDTMTabRS2("dat")) & vbTab & "" & Trim(selDTMTabRS2("iat"))
                        i = i + 1
                    Else
                        MSFrcs.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO")) & vbTab & "" & Trim(selDTMTabRS2("abo")) & vbTab & "" & Trim(selDTMTabRS2("rh")) & vbTab & "" & Trim(selDTMTabRS2("abs")) & vbTab & "" & Trim(selDTMTabRS2("dat")) & vbTab & "" & Trim(selDTMTabRS2("iat")) & vbTab & "" & SELDTMNO_RS("DTMNO")
                        i = i + 1
                    End If
                End If
                selDTMTabRS2.MoveNext
            Loop
        End If
    ElseIf cmb_option.Text = "HBsAg(ELISA)" Then
        SSTab1.TabVisible(0) = False
        SSTab1.TabVisible(1) = True
        SSTab1.TabVisible(2) = False
        SSTab1.TabVisible(3) = False
        MSFhbsAg.Rows = i
        If Not selDTMTabRS2.EOF Then
            Do While Not selDTMTabRS2.EOF
                If SELDMLPROCE_FLAG = False Then
                    SELDMLPROCE_QR(0) = selDTMTabRS2("REQNO")
                    Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                    SELDMLPROCE_FLAG = True
                Else
                    SELDMLPROCE_QR(0) = selDTMTabRS2("REQNO")
                    SELDMLPROCE_RS.MoreResults
                    SELDMLPROCE_RS.Requery
                End If
                proc = ""
                If Not SELDMLPROCE_RS.EOF Then
                    Do While Not SELDMLPROCE_RS.EOF
                        If SELDMLPROCE_RS("PRO_CODE") = "H012" Then
                            proc = "T"
                        End If
                        SELDMLPROCE_RS.MoveNext
                    Loop
                End If
                If proc = "T" Then
                    a = CInt(Trim(selDTMTabRS2("hbsagkey")))
                    If SELDTMNO_RS.EOF Then
                        If IsNull(selDTMTabRS2("hbsag")) Then
                            MSFhbsAg.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO"))
                        Else
                            MSFhbsAg.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO")) & vbTab & DECRPT(Trim(selDTMTabRS2("hbsag")), a)
                        End If
                        i = i + 1
                    Else
                        If IsNull(selDTMTabRS2("hbsag")) Then
                            MSFhbsAg.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO")) & vbTab & "" & vbTab & "" & SELDTMNO_RS("DTMNO")
                        Else
                            MSFhbsAg.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO")) & vbTab & DECRPT(Trim(selDTMTabRS2("hbsag")), a) & vbTab & "" & SELDTMNO_RS("DTMNO")
                        End If
                        i = i + 1
                    End If
                End If
                selDTMTabRS2.MoveNext
            Loop
        End If
    ElseIf cmb_option.Text = "HCV(ELISA)" Then
        SSTab1.TabVisible(0) = False
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = True
        SSTab1.TabVisible(3) = False
        MSFhcv.Rows = i
        If Not selDTMTabRS2.EOF Then
            Do While Not selDTMTabRS2.EOF
                If SELDMLPROCE_FLAG = False Then
                    SELDMLPROCE_QR(0) = selDTMTabRS2("REQNO")
                    Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                    SELDMLPROCE_FLAG = True
                Else
                    SELDMLPROCE_QR(0) = selDTMTabRS2("REQNO")
                    SELDMLPROCE_RS.MoreResults
                    SELDMLPROCE_RS.Requery
                End If
                proc = ""
                If Not SELDMLPROCE_RS.EOF Then
                    Do While Not SELDMLPROCE_RS.EOF
                        If SELDMLPROCE_RS("PRO_CODE") = "H014" Then
                            proc = "T"
                        End If
                        SELDMLPROCE_RS.MoveNext
                    Loop
                End If
                If proc = "T" Then
                    a = CInt(Trim(selDTMTabRS2("hcvkey")))
                    If SELDTMNO_RS.EOF Then
                        If IsNull(selDTMTabRS2("hcv")) Then
                            MSFhcv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO"))
                        Else
                            MSFhcv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO")) & vbTab & DECRPT(Trim(selDTMTabRS2("hcv")), a)
                        End If
                        i = i + 1
                    Else
                        If IsNull(selDTMTabRS2("hcv")) Then
                            MSFhcv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO")) & vbTab & "" & vbTab & "" & SELDTMNO_RS("DTMNO")
                        Else
                            MSFhcv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO")) & vbTab & DECRPT(Trim(selDTMTabRS2("hcv")), a) & vbTab & "" & SELDTMNO_RS("DTMNO")
                        End If
                        i = i + 1
                    End If
                End If
                selDTMTabRS2.MoveNext
            Loop
        End If
    ElseIf cmb_option.Text = "HIV(ELISA)" Then
        SSTab1.TabVisible(0) = False
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False
        SSTab1.TabVisible(3) = True
        MSFhiv.Rows = i
        If Not selDTMTabRS2.EOF Then
            Do While Not selDTMTabRS2.EOF
                If SELDMLPROCE_FLAG = False Then
                    SELDMLPROCE_QR(0) = selDTMTabRS2("REQNO")
                    Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                    SELDMLPROCE_FLAG = True
                Else
                    SELDMLPROCE_QR(0) = selDTMTabRS2("REQNO")
                    SELDMLPROCE_RS.MoreResults
                    SELDMLPROCE_RS.Requery
                End If
                proc = ""
                If Not SELDMLPROCE_RS.EOF Then
                    Do While Not SELDMLPROCE_RS.EOF
                        If SELDMLPROCE_RS("PRO_CODE") = "H013" Then
                            proc = "T"
                        End If
                        SELDMLPROCE_RS.MoveNext
                    Loop
                End If
                If proc = "T" Then
                    a = CInt(Trim(selDTMTabRS2("hivkey")))
                    If SELDTMNO_RS.EOF Then
                        If IsNull(selDTMTabRS2("hiv")) Then
                            MSFhiv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO"))
                        Else
                            MSFhiv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO")) & vbTab & DECRPT(Trim(selDTMTabRS2("hiv")), a)
                        End If
                        i = i + 1
                    Else
                        If IsNull(selDTMTabRS2("hiv")) Then
                            MSFhiv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO")) & vbTab & "" & vbTab & "" & SELDTMNO_RS("DTMNO")
                        Else
                            MSFhiv.AddItem i & vbTab & Trim(req_chglen(selDTMTabRS2("REQNO"))) & vbTab & Trim(selDTMTabRS2("CASENO")) & vbTab & DECRPT(Trim(selDTMTabRS2("hiv")), a) & vbTab & "" & SELDTMNO_RS("DTMNO")
                        End If
                        i = i + 1
                    End If
                End If
                selDTMTabRS2.MoveNext
            Loop
        End If
    End If
End If
End Sub
Private Sub BLANKALL()
Dim proc As String
Dim i
i = 1
If cmb_option.Text = "RCS" Then
    SSTab1.TabVisible(0) = True
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    If Drequisitn_FLAG_RCS_BLANK = False Then
       Drequisitn_QR_RCS_BLANK(0) = msk_fromdt.FormattedText
       Drequisitn_QR_RCS_BLANK(1) = msk_todt.FormattedText
       Set Drequisitn_RS_RCS_BLANK = Drequisitn_QR_RCS_BLANK.OpenResultset(1, 3)
       Drequisitn_FLAG_RCS_BLANK = True
    Else
       Drequisitn_QR_RCS_BLANK(0) = msk_fromdt.FormattedText
       Drequisitn_QR_RCS_BLANK(1) = msk_todt.FormattedText
       Drequisitn_RS_RCS_BLANK.MoreResults
       Drequisitn_RS_RCS_BLANK.Requery
    End If
    MSFrcs.Rows = i
    If Not Drequisitn_RS_RCS_BLANK.EOF Then
        Do While Not Drequisitn_RS_RCS_BLANK.EOF
            If selDTMTabFLAG1 = False Then
               selDTMTabQR1(0) = Trim(Drequisitn_RS_RCS_BLANK("REQNO"))
               Set selDTMTabRS1 = selDTMTabQR1.OpenResultset(1, 3)
               selDTMTabFLAG1 = True
            Else
                selDTMTabQR1(0) = Trim(Drequisitn_RS_RCS_BLANK("REQNO"))
                selDTMTabRS1.MoreResults
                selDTMTabRS1.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS_RCS_BLANK("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS_RCS_BLANK("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = Drequisitn_RS_RCS_BLANK("reqno")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = Drequisitn_RS_RCS_BLANK("reqno")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H001" Or SELDMLPROCE_RS("PRO_CODE") = "H007" Or SELDMLPROCE_RS("PRO_CODE") = "H006" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                If SELDTMNO_RS.EOF Then
                    MSFrcs.AddItem i & vbTab & req_chglen(Drequisitn_RS_RCS_BLANK("reqno")) & vbTab & Drequisitn_RS_RCS_BLANK("CASENO") & vbTab & "" & Trim(selDTMTabRS1("abo")) & vbTab & "" & Trim(selDTMTabRS1("rh")) & vbTab & "" & Trim(selDTMTabRS1("abs")) & vbTab & "" & Trim(selDTMTabRS1("dat")) & vbTab & "" & Trim(selDTMTabRS1("iat"))
                    i = i + 1
                Else
                    MSFrcs.AddItem i & vbTab & req_chglen(Drequisitn_RS_RCS_BLANK("reqno")) & vbTab & Drequisitn_RS_RCS_BLANK("CASENO") & vbTab & "" & Trim(selDTMTabRS1("abo")) & vbTab & "" & Trim(selDTMTabRS1("rh")) & vbTab & "" & Trim(selDTMTabRS1("abs")) & vbTab & "" & Trim(selDTMTabRS1("dat")) & vbTab & "" & Trim(selDTMTabRS1("iat")) & vbTab & " " & Trim(SELDTMNO_RS("DTMNO"))
                    i = i + 1
                End If
            End If
            Drequisitn_RS_RCS_BLANK.MoveNext
        Loop
    End If
ElseIf cmb_option.Text = "HBsAg(ELISA)" Then
    SSTab1.TabVisible(0) = False
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = False
    If Drequisitn_FLAG_HBSAG_BLANK = False Then
       Drequisitn_QR_HBSAG_BLANK(0) = msk_fromdt.FormattedText
       Drequisitn_QR_HBSAG_BLANK(1) = msk_todt.FormattedText
       Set Drequisitn_RS_HBSAG_BLANK = Drequisitn_QR_HBSAG_BLANK.OpenResultset(1, 3)
       Drequisitn_FLAG_HBSAG_BLANK = True
    Else
       Drequisitn_QR_HBSAG_BLANK(0) = msk_fromdt.FormattedText
       Drequisitn_QR_HBSAG_BLANK(1) = msk_todt.FormattedText
       Drequisitn_RS_HBSAG_BLANK.MoreResults
       Drequisitn_RS_HBSAG_BLANK.Requery
    End If
    MSFhbsAg.Rows = i
    If Not Drequisitn_RS_HBSAG_BLANK.EOF Then
        Do While Not Drequisitn_RS_HBSAG_BLANK.EOF
            If selDTMTabFLAG1 = False Then
               selDTMTabQR1(0) = Trim(Drequisitn_RS_HBSAG_BLANK("REQNO"))
               Set selDTMTabRS1 = selDTMTabQR1.OpenResultset(1, 3)
               selDTMTabFLAG1 = True
            Else
                selDTMTabQR1(0) = Trim(Drequisitn_RS_HBSAG_BLANK("REQNO"))
                selDTMTabRS1.MoreResults
                selDTMTabRS1.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS_HBSAG_BLANK("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS_HBSAG_BLANK("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = Drequisitn_RS_HBSAG_BLANK("reqno")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = Drequisitn_RS_HBSAG_BLANK("reqno")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H012" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                If SELDTMNO_RS.EOF Then
                    MSFhbsAg.AddItem i & vbTab & req_chglen(Drequisitn_RS_HBSAG_BLANK("REQNO")) & vbTab & Drequisitn_RS_HBSAG_BLANK("CASENO") & vbTab & "" & selDTMTabRS1("hbsag")
                    i = i + 1
                Else
                    MSFhbsAg.AddItem i & vbTab & req_chglen(Drequisitn_RS_HBSAG_BLANK("REQNO")) & vbTab & Drequisitn_RS_HBSAG_BLANK("CASENO") & vbTab & "" & selDTMTabRS1("hbsag") & vbTab & "" & SELDTMNO_RS("DTMNO")
                    i = i + 1
                End If
            End If
            Drequisitn_RS_HBSAG_BLANK.MoveNext
        Loop
    End If
ElseIf cmb_option.Text = "HCV(ELISA)" Then
    SSTab1.TabVisible(0) = False
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = True
    SSTab1.TabVisible(3) = False
    If Drequisitn_FLAG_HCV_BLANK = False Then
       Drequisitn_QR_HCV_BLANK(0) = msk_fromdt.FormattedText
       Drequisitn_QR_HCV_BLANK(1) = msk_todt.FormattedText
       Set Drequisitn_RS_HCV_BLANK = Drequisitn_QR_HCV_BLANK.OpenResultset(1, 3)
       Drequisitn_FLAG_HCV_BLANK = True
    Else
       Drequisitn_QR_HCV_BLANK(0) = msk_fromdt.FormattedText
       Drequisitn_QR_HCV_BLANK(1) = msk_todt.FormattedText
       Drequisitn_RS_HCV_BLANK.MoreResults
       Drequisitn_RS_HCV_BLANK.Requery
    End If
    MSFhcv.Rows = i
    If Not Drequisitn_RS_HCV_BLANK.EOF Then
        Do While Not Drequisitn_RS_HCV_BLANK.EOF
            If selDTMTabFLAG1 = False Then
               selDTMTabQR1(0) = Trim(Drequisitn_RS_HCV_BLANK("REQNO"))
               Set selDTMTabRS1 = selDTMTabQR1.OpenResultset(1, 3)
               selDTMTabFLAG1 = True
            Else
                selDTMTabQR1(0) = Trim(Drequisitn_RS_HCV_BLANK("REQNO"))
                selDTMTabRS1.MoreResults
                selDTMTabRS1.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS_HCV_BLANK("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS_HCV_BLANK("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = Drequisitn_RS_HCV_BLANK("reqno")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = Drequisitn_RS_HCV_BLANK("reqno")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H014" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                If SELDTMNO_RS.EOF Then
                    MSFhcv.AddItem i & vbTab & req_chglen(Drequisitn_RS_HCV_BLANK("REQNO")) & vbTab & Drequisitn_RS_HCV_BLANK("CASENO") & vbTab & "" & selDTMTabRS1("hcv")
                    i = i + 1
                Else
                    MSFhcv.AddItem i & vbTab & req_chglen(Drequisitn_RS_HCV_BLANK("REQNO")) & vbTab & Drequisitn_RS_HCV_BLANK("CASENO") & vbTab & "" & selDTMTabRS1("hcv") & vbTab & SELDTMNO_RS("DTMNO")
                    i = i + 1
                End If
            End If
            Drequisitn_RS_HCV_BLANK.MoveNext
        Loop
    End If
ElseIf cmb_option.Text = "HIV(ELISA)" Then
    SSTab1.TabVisible(0) = False
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    SSTab1.TabVisible(3) = True
    If Drequisitn_FLAG_HIV_BLANK = False Then
       Drequisitn_QR_HIV_BLANK(0) = msk_fromdt.FormattedText
       Drequisitn_QR_HIV_BLANK(1) = msk_todt.FormattedText
       Set Drequisitn_RS_HIV_BLANK = Drequisitn_QR_HIV_BLANK.OpenResultset(1, 3)
       Drequisitn_FLAG_HIV_BLANK = True
    Else
       Drequisitn_QR_HIV_BLANK(0) = msk_fromdt.FormattedText
       Drequisitn_QR_HIV_BLANK(1) = msk_todt.FormattedText
       Drequisitn_RS_HIV_BLANK.MoreResults
       Drequisitn_RS_HIV_BLANK.Requery
    End If
    MSFhiv.Rows = i
    If Not Drequisitn_RS_HIV_BLANK.EOF Then
        Do While Not Drequisitn_RS_HIV_BLANK.EOF
            If selDTMTabFLAG1 = False Then
               selDTMTabQR1(0) = Trim(Drequisitn_RS_HIV_BLANK("REQNO"))
               Set selDTMTabRS1 = selDTMTabQR1.OpenResultset(1, 3)
               selDTMTabFLAG1 = True
            Else
                selDTMTabQR1(0) = Trim(Drequisitn_RS_HIV_BLANK("REQNO"))
                selDTMTabRS1.MoreResults
                selDTMTabRS1.Requery
            End If
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = Drequisitn_RS_HIV_BLANK("CASENO")
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = Drequisitn_RS_HIV_BLANK("CASENO")
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELDMLPROCE_FLAG = False Then
                SELDMLPROCE_QR(0) = Drequisitn_RS_HIV_BLANK("reqno")
                Set SELDMLPROCE_RS = SELDMLPROCE_QR.OpenResultset(1, 3)
                SELDMLPROCE_FLAG = True
            Else
                SELDMLPROCE_QR(0) = Drequisitn_RS_HIV_BLANK("reqno")
                SELDMLPROCE_RS.MoreResults
                SELDMLPROCE_RS.Requery
            End If
            proc = ""
            If Not SELDMLPROCE_RS.EOF Then
                Do While Not SELDMLPROCE_RS.EOF
                    If SELDMLPROCE_RS("PRO_CODE") = "H013" Then
                        proc = "T"
                    End If
                    SELDMLPROCE_RS.MoveNext
                Loop
            End If
            If proc = "T" Then
                If SELDTMNO_RS.EOF Then
                    MSFhiv.AddItem i & vbTab & req_chglen(Drequisitn_RS_HIV_BLANK("REQNO")) & vbTab & Drequisitn_RS_HIV_BLANK("CASENO") & vbTab & "" & Trim(selDTMTabRS1("hiv"))
                    i = i + 1
                Else
                    MSFhiv.AddItem i & vbTab & req_chglen(Drequisitn_RS_HIV_BLANK("REQNO")) & vbTab & Drequisitn_RS_HIV_BLANK("CASENO") & vbTab & "" & Trim(selDTMTabRS1("hiv")) & vbTab & "" & SELDTMNO_RS("DTMNO")
                    i = i + 1
                End If
            End If
            Drequisitn_RS_HIV_BLANK.MoveNext
        Loop
    End If
End If
End Sub
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



