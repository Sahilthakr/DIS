VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tmh_ris_appointment 
   Caption         =   "Appointment "
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   15
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   13815
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Cr_appointment 
      Left            =   0
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame frm_cmd 
      Height          =   1155
      Left            =   150
      TabIndex        =   27
      Top             =   7440
      Width           =   11535
      Begin VB.CommandButton cmd_query 
         Caption         =   "Query"
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
         Left            =   10080
         TabIndex        =   65
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_exit 
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
         Left            =   8640
         TabIndex        =   31
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
         Left            =   7320
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_delete 
         Caption         =   "&Delete"
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
         Left            =   2880
         TabIndex        =   29
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
         Left            =   1560
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_new 
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
         TabIndex        =   0
         Top             =   240
         Width           =   975
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
         Left            =   5760
         TabIndex        =   32
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
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_msg 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   47
         Top             =   720
         Width           =   8175
      End
      Begin VB.Label Label22 
         Caption         =   "Message Box "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   33
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame frm_entry 
      Height          =   3015
      Left            =   120
      TabIndex        =   23
      Top             =   4380
      Width           =   11595
      Begin VB.CommandButton CMD_DIALY 
         Caption         =   "&Diary"
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
         TabIndex        =   66
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cbo_location 
         Height          =   315
         Left            =   3150
         TabIndex        =   4
         Top             =   480
         Width           =   1965
      End
      Begin VB.ComboBox cbo_year 
         Height          =   315
         Left            =   10320
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cbo_month 
         Height          =   315
         ItemData        =   "tmh_ris_appointment.frx":0000
         Left            =   8700
         List            =   "tmh_ris_appointment.frx":0002
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmd_help 
         Caption         =   "&Help"
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
         TabIndex        =   10
         Top             =   1980
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_instructions 
         Height          =   975
         Left            =   1590
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1740
         Width           =   3525
      End
      Begin VB.ComboBox cbo_apttm 
         Height          =   315
         ItemData        =   "tmh_ris_appointment.frx":0004
         Left            =   840
         List            =   "tmh_ris_appointment.frx":0050
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1110
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid msf_calendar 
         Height          =   1815
         Left            =   7620
         TabIndex        =   8
         Top             =   840
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   7
         Cols            =   7
         FixedCols       =   0
         GridColor       =   8421504
         ScrollBars      =   0
         FormatString    =   "SUN | MON | TUE | WED | THU | FRI | SAT"
      End
      Begin VB.Label Label20 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2250
         TabIndex        =   55
         Top             =   540
         Width           =   945
      End
      Begin VB.Label lbl_apttm 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   53
         Top             =   1170
         Width           =   1335
      End
      Begin VB.Label lbl_monthyear 
         Height          =   375
         Left            =   4680
         TabIndex        =   52
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lbl_aptdt 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   51
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "Instructions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   26
         Top             =   1740
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1110
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   540
         Width           =   555
      End
   End
   Begin VB.Frame frm_disp 
      Height          =   2235
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   11595
      Begin MSMask.MaskEdBox msk_lmpdate 
         Height          =   315
         Left            =   10020
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label LBL_LMP 
         Caption         =   "LMP Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         TabIndex        =   63
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Bed No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9480
         TabIndex        =   57
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lbl_bedno 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   10320
         TabIndex        =   56
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lbl_patregtype 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7440
         TabIndex        =   54
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbl_proccode 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   10320
         TabIndex        =   46
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lbl_procname 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2220
         TabIndex        =   45
         Top             =   1800
         Width           =   2955
      End
      Begin VB.Label Label16 
         Caption         =   "Name of Procedure"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1800
         Width           =   2115
      End
      Begin VB.Label Label17 
         Caption         =   "Procedure Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8580
         TabIndex        =   43
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label lbl_ipopsts 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   10320
         TabIndex        =   40
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl_wrdbdno 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6360
         TabIndex        =   39
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lbl_refby 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2220
         TabIndex        =   38
         Top             =   1020
         Width           =   3015
      End
      Begin VB.Label lbl_catg 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7920
         TabIndex        =   37
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbl_age 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4620
         TabIndex        =   36
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lbl_sex 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2220
         TabIndex        =   35
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbl_name 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2220
         TabIndex        =   34
         Top             =   180
         Width           =   4575
      End
      Begin VB.Label Label13 
         Caption         =   "Ward :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   22
         Top             =   1020
         Width           =   555
      End
      Begin VB.Label Label12 
         Caption         =   "Primary Doctor/Service"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   540
         TabIndex        =   21
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label11 
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6420
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   19
         Top             =   660
         Width           =   675
      End
      Begin VB.Label Label9 
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4020
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1710
         TabIndex        =   17
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label7 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1470
         TabIndex        =   16
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Frame frm_key 
      Height          =   1275
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   11595
      Begin VB.ComboBox cbo_deptname 
         Height          =   315
         ItemData        =   "tmh_ris_appointment.frx":00F8
         Left            =   2220
         List            =   "tmh_ris_appointment.frx":00FA
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   180
         Width           =   2535
      End
      Begin MSMask.MaskEdBox msk_caseno 
         Height          =   315
         Left            =   2220
         TabIndex        =   2
         ToolTipText     =   "ENTER THE CASE NUMBER TO DISPLAY PATIENT DETAILS"
         Top             =   480
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cbo_reqno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "tmh_ris_appointment.frx":00FC
         Left            =   2220
         List            =   "tmh_ris_appointment.frx":00FE
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   58
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lbl_modalitycode 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9540
         TabIndex        =   50
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label Label21 
         Caption         =   "Modality Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   49
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label lbl_modality 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9540
         TabIndex        =   48
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Modality"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8520
         TabIndex        =   42
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "Requisition No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   540
         TabIndex        =   14
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Case No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1140
         TabIndex        =   13
         Top             =   540
         Width           =   975
      End
   End
   Begin MSMask.MaskEdBox msk_sysdate 
      Height          =   315
      Left            =   10440
      TabIndex        =   59
      Top             =   180
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox msk_systime 
      Height          =   315
      Left            =   10560
      TabIndex        =   60
      Top             =   540
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm:ss"
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   67
      Top             =   0
      Width           =   9375
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9720
      TabIndex        =   62
      Top             =   270
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9660
      TabIndex        =   61
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label30 
      Caption         =   "APPOINTMENT"
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
      Height          =   255
      Left            =   4380
      TabIndex        =   41
      Top             =   420
      Width           =   2265
   End
End
Attribute VB_Name = "tmh_ris_appointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer, j As Integer, K As Integer, totholidays As Integer, NVLROWS As Integer
Dim D1 As String, D2 As String
Dim todt() As String, fromdt() As String
Dim m As Integer, Y As Integer
Dim hlist() As String
Dim totnonavldays As Integer
Dim srvnonavail_list() As String
Dim RS As rdoResultset
Dim zpublicHoliday As Boolean
'Dim QRYFLAG As Boolean
Dim CASENUM As String
Dim REQNUM As String
Dim MSK_REQNUM As String
Dim LONGMSK_REQNUM As String
'******from req fh******8

Dim SELMA101_QR As rdoQuery
Dim SELMA101_RS As rdoResultset

Dim rq_dept As rdoQuery
Dim rs_dept As rdoResultset
Dim flg_dept As Boolean '**5/7/2001

Dim rq_deptcode As rdoQuery
Dim rs_deptcode As rdoResultset

Dim rq_patdet As rdoQuery
Dim rs_patdet As rdoResultset
Dim flg_patdet As Boolean

Dim rq_patipop As rdoQuery
Dim rs_patipop As rdoResultset
Dim flg_patipop As Boolean

Dim rq_modality As rdoQuery
Dim rs_modality As rdoResultset

Dim rq_moddes As rdoQuery
Dim rs_moddes As rdoResultset
Dim flg_moddes As Boolean


Dim rq_refdoc As rdoQuery
Dim rs_refdoc As rdoResultset

Dim rq_proc As rdoQuery
Dim rs_proc As rdoResultset

Dim rq_procname As rdoQuery
Dim rs_procname As rdoResultset
Dim flg_procname As Boolean

Dim rq_provdiag As rdoQuery
Dim rs_provdiag As rdoResultset

Dim RREQUISITNInsrt_Qr As rdoQuery

Dim rq_RREQUISITN As rdoQuery
Dim rs_RREQUISITN As rdoResultset
Dim flg_RREQUISITN As Boolean

Dim new_view_flg As Integer
Dim str As String  '***********till req th******

Dim rq_reqpend As rdoQuery
Dim rs_reqpend As rdoResultset
Dim flg_reqpend As Boolean

'******add 29052001 from here
Dim rq_apptgiven As rdoQuery
Dim rs_apptgiven As rdoResultset
Dim flg_apptgiven As Boolean
'*****29052001 till here

Dim rq_serinstrct As rdoQuery
Dim rs_serinstrct As rdoResultset
Dim flg_serinstrct As Boolean

Dim rq_appinstr As rdoQuery
Dim rs_appinstr As rdoResultset
Dim flg_appinstr As Boolean

'****add 5/7/2001**********
Dim rq_applocation As rdoQuery
Dim rs_applocation As rdoResultset
Dim flg_applocation As Boolean
'*************

Dim rq_RREQUISITNApptInsrt_Qr As rdoQuery
Dim rq_RREQUISITNApptDel_Qr As rdoQuery
Dim rq_RREQUISITNApptUpdt_Qr As rdoQuery
Dim rq_AppInstrInsrt_Qr As rdoQuery
Dim rq_AppInstrDel_Qr As rdoQuery

Dim rq_calfill As rdoQuery
Dim rs_calfill As rdoResultset
Dim flg_calfill As Boolean

Dim rq_sernonavail As rdoQuery
Dim rs_sernonavail As rdoResultset
Dim flg_sernonavail As Boolean

Dim rq_clinhistlmpdt As rdoQuery    '*******add 15/09/2001
Dim rs_clinhistlmpdt As rdoResultset
Dim flg_clinhistlmpdt As Boolean

Dim calvar_holiday As Integer
Dim reqnextstg_flg As Integer
Dim MEMOGEN_VAR As Integer '*******4/9/2001
Dim var_aptnotgiven As Integer
Dim var_txtinstruct As String
Dim var_nopendingreq As Integer
Dim var_olddeptname As String '*******5/7/2001
Dim var_cbodepttxt As String '*****11/7/2001 (used for stopping the deptname and location combo getting cleared when form being activated after seeing the appt slot form.)
Dim LMPDT_VIS As Integer '***15/9/2001

Dim rq_caseapptgiven As rdoQuery '**15/09/2001(for checking if the same case number already given appts on the same day")
Dim rs_caseapptgiven As rdoResultset
Dim flg_caseapptgiven As Boolean
Dim Insrt_Stmt As rdoQuery
Dim selrislnklf_qr As rdoQuery
Dim selrislnklf_rs As rdoResultset

Dim delta102_qr As rdoQuery
Dim No_Chargeflag As Boolean
Dim VAR_DCD As String '**deptcode
Dim RISAUTH_QR As rdoQuery
Dim RISAUTH_RS As rdoResultset
Dim refdate
Dim refdatef
Dim refdatet
Dim refdateft
Dim refdatett
Dim selreq_qr As rdoQuery
Dim selreq_rs As rdoResultset
Dim strtime
Dim timeslotno  As Integer

Public Sub NO_CHARGE()
On Error GoTo MSG:
No_Chargeflag = False

SELMA101_QR(0) = Trim(lbl_proccode.Caption)
SELMA101_RS.MoreResults
SELMA101_RS.Requery
If SELMA101_RS.RowCount > 0 Then
    If Trim(lbl_catg.Caption) = "C" Then
        If SELMA101_RS("C_CH") = 0 Then
            No_Chargeflag = True
        End If
    ElseIf Trim(lbl_catg.Caption) = "NC" Then
        If SELMA101_RS("NC_CH") = 0 Then
            No_Chargeflag = True
        End If
    End If
Else
    MsgBox "There is no Record with this SERVICE CODE."
End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub




Sub VALID_DT()
On Error GoTo MSG:
Dim nday1 As String
Dim nday2
Dim i As Integer
Dim daterr

On Error GoTo msgerr:

daterr = False
Dim aptdt As String
aptdt = lbl_aptdt.Caption
calvar_holiday = 0
If IsDate(aptdt) Then

'If IsDate(Format(lbl_aptdt.Caption, "##/##/####")) Then
    nday2 = Right(lbl_aptdt.Caption, 4) & "-" & Mid(lbl_aptdt.Caption, 4, 2) & "-" & Left(lbl_aptdt.Caption, 2)
        'If nday2 < Format(Date, "yyyy-mm-dd") Then '*** FUTURE DATES ONLY'11599
       'If nday2 < GET_REVDATE(msk_sysdate.FormattedText) Then '*** FUTURE DATES ONLY'11599
       If GET_REVDATE(lbl_aptdt.Caption) < GET_REVDATE(msk_sysdate.FormattedText) Then '*** FUTURE DATES ONLY'11599

        MsgBox "Please give only Future DATES for Appointments."
        calvar_holiday = 1
        Exit Sub
    End If

    nday1 = Weekday(Format(nday2))  '*** SAT AND SUN DAYS ONLY
    'If nday1 = 1 Or nday1 = 7 Then'****comm 29052001
   If nday1 = 7 Then
   MsgBox "This is a SATARDAY and ONLY LIMITED appointments are to be given."
   End If
    If nday1 = 1 Then   '****only sundays'****add 29052001
        'i = MsgBox("No appointments on saturdays and sundays!" & Chr(13) & " Do you still wish to issue appointment?", vbYesNo)
        i = MsgBox("No appointments on sundays!" & Chr(13) & " Do you still wish to issue appointment?", vbYesNo)

        If i = vbNo Then
              calvar_holiday = 1
              daterr = True
              Exit Sub
        End If
    Else                    '*** NOT HOLIDAYS
    Dim rs9 As rdoResultset
    'Set RS_aPP = rdb.OpenResultset("SELECT * FROM RREQUISITN WHERE CSTR(APNTDT) = '" & Format(lbl_aptdate.Caption, "MM/DD/YYYY") & "'", 1, 1)   ' + Format(lbl_aptdate.Caption, "MM/DD/YYYY"), 1, 1)
    'Set rs9 = rdb.OpenResultset("SELECT hday FROM PATOOLSLIB.ma1121 WHERE CSTR(hday) = '" & Format(lbl_aptdt.Caption, "dd/mm/YYYY") & "'", 1, 1)   ' + Format(lbl_aptdate.Caption, "MM/DD/YYYY"), 1, 1)
   ' Set rs9 = rdb.OpenResultset("SELECT hday FROM PATOOLSLIB.ma1121 WHERE hday = '" & Format(lbl_aptdt.Caption, "YYYY-MM-DD") & "'", 1, 1)   ' + Format(lbl_aptdate.Caption, "MM/DD/YYYY"), 1, 1)
    Set rs9 = rdb.OpenResultset("SELECT hday FROM PATOOLSLIB.ma1121 WHERE hday = '" & GET_REVDATE(lbl_aptdt.Caption) & "'", 1, 1)   ' + Format(lbl_aptdate.Caption, "MM/DD/YYYY"), 1, 1)

If rs9.RowCount > 0 Then
        rs9.MoveFirst
        'If rs1.RowCount > 0 Then
        While Not rs9.EOF
           'MsgBox Format(rs9("hday"), "YYYY-MM-DD")
            'If Format(rs9("hday"), "YYYY-MM-DD") = nday2 Then
            Dim dt1a
            Dim dt2a
            dt1a = GET_REVDATE(Format(rs9("hday"), "dd/mm/yyyy"))
            dt2a = GET_REVDATE(Format(nday2, "dd/mm/yyyy"))
            If dt1a = dt2a Then
               '25/01/2002-Parag
               'Boolean Value set to true if Selected date matches with Public Holiday Date
               zpublicHoliday = True
               i = MsgBox("This is a HOLIDAY,so DO NOT give ANY APPOINTMENTS.", vbCritical)
               calvar_holiday = 1
            End If
            rs9.MoveNext
        Wend
        End If
    End If
  '*** NOT ON LEAVES
''Set rs1 = rdb.OpenResultset("select * from ma1122 where doc_name = '" & c_referred_to.Text & "'", 2, 3)
'If rs10.RowCount > 0 Then
'rs10.MoveFirst
'    While Not rs10.EOF
'            If rs10("doc_name") = c_referred_to.Text Then
'                a1 = DateDiff("d", nday2, rs10("FROM_date"))
'                a2 = DateDiff("d", nday2, rs10("TO_date"))
'
'                If a1 <= 0 Then
'                    If a2 >= 0 Then
'                        'MsgBox "Doctor is on leave - No appointment can be given"
'                        i = MsgBox("Doctor is on leave!" & Chr(13) & "Do you still wish to issue appointment?", vbYesNo)
'                        If i = vbNo Then '**4/3/99
'                            me_rdate.SetFocus
'                            daterr = True
'                            Exit Sub
'                        End If '**4/3/99
'                    End If
'                End If
'            End If
'              rs10.MoveNext
'    Wend
'End If
Else
    MsgBox "The Appointment Date is NOT correct."
'    If me_rdate.Visible = True Then
'        me_rdate.SetFocus
'    ElseIf Text5.Visible = True And Text5.Enabled = True Then
'        Text5.SetFocus
'    End If
    calvar_holiday = 1
    daterr = True
    Exit Sub
End If

Exit Sub

msgerr: MsgBox Error(Err)
Exit Sub
MSG:
MsgBox Error(Err)

End Sub


Sub setflagsfalse()
    flg_moddes = False
    flg_RREQUISITN = False
    flg_procname = False
    flg_reqpend = False
    flg_patdet = False
    flg_patipop = False
    flg_serinstrct = False
    flg_appinstr = False
    flg_calfill = False
    flg_sernonavail = False
    flg_apptgiven = False '****29052001
    flg_applocation = False '***5/7/2001
    flg_dept = False '****5/7/2001
    flg_clinhistlmpdt = False '**15092001
    flg_caseapptgiven = False '**15092001
End Sub



Sub cmdbtn_vis(n As String)

If n = "I" Then

'    cmdnew
'    cmdview
 CMD_NEW.Enabled = True
 cmd_view.Enabled = True
 cmd_exit.Enabled = True
 'cmd_modify.Enabled = False
 cmd_delete.Enabled = False
 cmd_print.Enabled = False
 cmd_save.Enabled = False
 cmd_cancel.Enabled = False

 ElseIf n = "N" Then

    'cmdsave
    cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_exit.Enabled = True
     CMD_NEW.Enabled = False
    cmd_view.Enabled = False
    'cmd_modify.Enabled = False
    cmd_delete.Enabled = False
    cmd_print.Enabled = False




ElseIf n = "V" Then
cmd_save.Enabled = False
CMD_NEW.Enabled = False
cmd_view.Enabled = False
'cmdmodify
'CMDDELETE
cmd_delete.Enabled = True
'CMDPRINT
cmd_print.Enabled = True
cmd_cancel.Enabled = True
cmd_exit.Enabled = True


ElseIf n = "M" Then
' cmdnew
' cmdview
' CMDDELETE
' cmdsave
' CMDPRINT
' CMDQUERY
' CMDHELP
' CMDDIARY

'cmd_save.Enabled = True
'cmd_cancel.Enabled = True
'cmd_exit.Enabled = True
'
'cmd_new.Enabled = False
'cmd_view.Enabled = False
''cmd_modify.Enabled = False
'cmd_delete.Enabled = False
'cmd_print.Enabled = False
''cmd_commit.Enabled = False
ElseIf n = "D" Then
    'cmdnew
    'cmdview
    CMD_NEW.Enabled = True
    cmd_view.Enabled = True
    cmd_save.Enabled = False
    cmd_cancel.Enabled = True
    'cmd_modify.Enabled = False
    cmd_delete.Enabled = False
    cmd_print.Enabled = False
    cmd_exit.Enabled = True

ElseIf n = "S" Then
' cmdnew
' cmdview
' CMDPRINT
 CMD_NEW.Enabled = True
 cmd_view.Enabled = True
 cmd_print.Enabled = True
cmd_save.Enabled = False
cmd_cancel.Enabled = True
'cmd_modify.Enabled = False
cmd_delete.Enabled = False
cmd_exit.Enabled = True

ElseIf n = "CA" Then
 'cmdnew
 'cmdview
 CMD_NEW.Enabled = True
 cmd_view.Enabled = True
cmd_exit.Enabled = True
'cmd_modify.Enabled = False
cmd_delete.Enabled = False
cmd_print.Enabled = False
cmd_save.Enabled = False
cmd_cancel.Enabled = True


ElseIf n = "CO" Then
'    cmdnew
'    cmdview
'    CMDPRINT
CMD_NEW.Enabled = True
 cmd_view.Enabled = True
 cmd_print.Enabled = True
cmd_exit.Enabled = True
'cmd_modify.Enabled = False
cmd_delete.Enabled = False
cmd_save.Enabled = False
cmd_cancel.Enabled = False

ElseIf n = "P" Then



cmd_exit.Enabled = True

'cmd_modify.Enabled = False
cmd_delete.Enabled = False
cmd_print.Enabled = False
cmd_save.Enabled = False
cmd_cancel.Enabled = True
 CMD_NEW.Enabled = True
 cmd_view.Enabled = True
 
' cmdnew
' cmdview


ElseIf n = "YES" Then
 'cmdnew
 'cmdview
CMD_NEW.Enabled = True
 cmd_view.Enabled = True
cmd_exit.Enabled = True
'
'cmd_modify.Enabled = False
cmd_delete.Enabled = False
cmd_print.Enabled = False
cmd_save.Enabled = False
cmd_cancel.Enabled = True


ElseIf n = "NO" Then
 'cmdnew
 'cmdview

CMD_NEW.Enabled = True
 cmd_view.Enabled = True
cmd_exit.Enabled = True

'cmd_modify.Enabled = False
cmd_delete.Enabled = False
cmd_print.Enabled = False
cmd_save.Enabled = False
cmd_cancel.Enabled = True

End If

End Sub




Private Sub cbo_deptname_Click()
'clear_fields
'cbo_deptname.Text = var_olddeptname
'lbl_msg.Caption = "select Department for appointment is to be given"

End Sub

Private Sub cbo_deptname_GotFocus()
lbl_msg.Caption = "select Department for appointment is to be given"

End Sub

Private Sub cbo_deptname_LostFocus()
On Error GoTo MSG:
If Len(Trim(cbo_deptname.Text)) <> 0 Then
    msk_caseno.Enabled = True
    msk_caseno.SetFocus
Else
    Dim j
    j = MsgBox("The Department Name is NOT selected.", vbCritical)
    If cbo_deptname.Enabled = True Then
    cbo_deptname.SetFocus
    End If

End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cbo_location_KeyPress(KeyAscii As Integer) '***added on 5/7/2001
If KeyAscii = 36 Or KeyAscii = 39 Or KeyAscii = 124 Then
KeyAscii = 0
End If

End Sub

'Private Sub Calendar1_Click()
'    lbl_aptdt.Caption = Format(Calendar1.Value, "dd/mm/yyyy")
'    VALID_DT
'    If calvar_holiday = 1 Then
'        lbl_aptdt.Caption = ""
'        Calendar1.SetFocus
'    End If
'End Sub

Private Sub cbo_reqno_Click()
On Error GoTo MSG:
    If new_view_flg = 1 Then '****29052001
        If Len(Trim(cbo_reqno.Text)) <> 0 Then
            REQNUM = req_chglen(cbo_reqno.Text)
            RREQUISITN_retrieve

            '*******txt changed to cbo location 5/7/2001****
            If cbo_location.Enabled = True Then '******'If txt_location.Enabled = True Then
                cbo_location.SetFocus '***'    txt_location.SetFocus
            End If
            '**************************
        End If
    End If '***add 29052001

    '*** 29052001 add from here
    If new_view_flg = 2 Then
     mskreplacedcbo
     'cmdbtn_vis ("V")
    End If
    '****29052001 add till here
'***SERVICE NON AVAILABILITY********marking in the calendar
    If new_view_flg = 1 Then
        f1
        If Trim(lbl_Modalitycode.Caption) = "2D" Or Trim(lbl_Modalitycode.Caption) = "ST" Or Trim(lbl_Modalitycode.Caption) = "PF" Then
        Else
            fillcal_sernonavail
        End If
        f2
    End If
    
'******
 '*********filling appt time combo**********
'    Dim strtime
'    Dim timeslotno  As Integer

    cbo_apttm.clear
    If Trim(lbl_Modalitycode.Caption) = "MM" Or Trim(lbl_Modalitycode.Caption) = "MR" Or Trim(lbl_Modalitycode.Caption) = "CT" Or Trim(lbl_Modalitycode.Caption) = "CR" Then
        timeslotno = 1
        strtime = 8
        While timeslotno < 25
            If strtime < 10 Then
            strtime = "0" & strtime
            End If
            cbo_apttm.AddItem strtime & ":00"
            cbo_apttm.AddItem strtime & ":30"
            strtime = Val(strtime) + 1
            timeslotno = timeslotno + 2
        Wend
    End If
    If Trim(lbl_Modalitycode.Caption) = "2D" Or Trim(lbl_Modalitycode.Caption) = "PF" Or Trim(lbl_Modalitycode.Caption) = "ST" Then
        timeslotno = 1
        strtime = 9
        While timeslotno < 20
            If strtime < 10 Then
            strtime = "0" & strtime
            End If
            cbo_apttm.AddItem strtime & ":00"
            cbo_apttm.AddItem strtime & ":20"
            cbo_apttm.AddItem strtime & ":40"
            strtime = Val(strtime) + 1
            timeslotno = timeslotno + 3
        Wend
    End If

    '********comparing each time slot in apttime combo with each element in tms_array********
   

    
'******
'If Trim(USRGRP) = "TYPISTS" Then
'selreq_qr(0) = Trim(GET_REVDATE(CStr(DATE11)))
'selreq_qr(1) = Trim(Mid(cbo_reqno.Text, 5, 7))
'selreq_rs.MoreResults
'selreq_rs.Requery
'If selreq_rs.RowCount > 0 Then
' If selreq_rs.RowCount = 20 Then
'    MsgBox "TYPIST CAN GIVE ONLY 20 APPOINTMENTS FOR CTSCAN AND US "
'    cmdbtn_vis ("I")
'    Exit Sub
' End If
' If selreq_rs.RowCount = 6 Then
'    MsgBox "TYPIST CAN GIVE ONLY 6 APPOINTMENTS FOR MR"
'    cmdbtn_vis ("I")
'    Exit Sub
' End If
''If selreq_rs.RowCount = 20 Then
''    MsgBox "TYPIST CAN GIVE ONLY 20 APPOINTMENTS FOR US"
''    cmdbtn_vis ("I")
''    Exit Sub
'' End If
'
'End If
'End If
  Me.MousePointer = 0
Exit Sub
MSG:
MsgBox Error(Err)
End Sub
Sub DEPTCODE_FIND()
On Error GoTo MSG:
    '********5/7/2001 for selecting dept.****
    Dim K As Integer
    ''''Dim VAR_DCD As String '**deptcode
    K = InStr(1, cbo_deptname.Text, "$")
    VAR_DCD = Trim(Mid(cbo_deptname, K + 1))
    '*************************************
Exit Sub
MSG:
MsgBox Error(Err)
End Sub



Sub RREQUISITN_retrieve()
On Error GoTo MSG:
On Error GoTo errlbl:

MEMOGEN_VAR = 0
var_aptnotgiven = 0
var_txtinstruct = ""
LONGMSK_REQNUM = req_chglen(MSK_REQNUM)


If new_view_flg = 1 Then
    rq_moddes(0) = Mid(cbo_reqno.Text, 5, 2)
Else
    rq_moddes(0) = Mid(LONGMSK_REQNUM, 5, 2)
End If

If flg_moddes = False Then
    Set rs_moddes = rq_moddes.OpenResultset(1, 1)
    flg_moddes = True
Else
    rs_moddes.MoreResults
    rs_moddes.Requery
End If
If rs_moddes.RowCount > 0 Then
    lbl_modality.Caption = rs_moddes(0)
End If

If new_view_flg = 1 Then
    lbl_Modalitycode = Mid(cbo_reqno.Text, 5, 2)
Else
    lbl_Modalitycode = Mid(LONGMSK_REQNUM, 5, 2)
End If



If new_view_flg = 1 Then
    rq_RREQUISITN(0) = req_chglen(cbo_reqno.Text)
Else
    rq_RREQUISITN(0) = MSK_REQNUM
End If

  Screen.MousePointer = 11
  lbl_msg.Caption = "retrieving requisition details..."
  DoEvents
If flg_RREQUISITN = False Then
    Set rs_RREQUISITN = rq_RREQUISITN.OpenResultset(1, 1)
    flg_RREQUISITN = True
Else
    rs_RREQUISITN.MoreResults
    rs_RREQUISITN.Requery
End If

'******comm 29052001 from here***********
'  If new_view_flg = 2 Then
'  If rs_RREQUISITN.RowCount = 0 Then
'    Dim i
'    i = MsgBox("Requisition number not existing", vbCritical)
'    Screen.MousePointer = 1
'    lbl_msg.Caption = ""
'    msk_REQNNO.SelStart = 10
'    msk_REQNNO.SelLength = 6
'    msk_REQNNO.SetFocus '**29052001
'        Exit Sub
'  End If
'  End If
'******comm 29052001 till here***********



If IsNull(rs_RREQUISITN("APNTDT")) = False Then
    lbl_aptdt.Caption = Format(rs_RREQUISITN("APNTDT"), "dd/mm/yyyy")
Else
    lbl_aptdt.Caption = ""
    If new_view_flg = 2 Then
        MsgBox "The Appointment is NOT given."
        var_aptnotgiven = 1
    End If
End If




If IsNull(rs_RREQUISITN("APNTTM")) = False Then
    If cbo_apttm.Visible = True Then
        lbl_apttm.Caption = rs_RREQUISITN("APNTTM")
    End If
    If lbl_apttm.Visible = True Then
        lbl_apttm.Caption = rs_RREQUISITN("APNTTM")
    End If
Else
    If lbl_apttm.Visible = True Then
        lbl_apttm.Caption = ""
    End If
    If cbo_apttm.Visible = True Then
         cbo_apttm.ListIndex = -1
    End If

    If new_view_flg = 2 Then
     If var_aptnotgiven <> 1 Then '***apt date specified but not time
        MsgBox "The Appointment time is NOT specified."
     End If
    End If
End If

'************FILLING LOCATION CBO DEPENDING UPON DEPT AND MODALITY****14/09/2001***

DEPTCODE_FIND


rq_applocation(0) = VAR_DCD
rq_applocation(1) = lbl_Modalitycode
If flg_applocation = False Then
    Set rs_applocation = rq_applocation.OpenResultset(1, 1)
    flg_applocation = True
Else
    rs_applocation.MoreResults
    rs_applocation.Requery
End If

cbo_location.clear
If rs_applocation.RowCount > 0 Then
    While Not rs_applocation.EOF
        cbo_location.AddItem rs_applocation("location")
        rs_applocation.MoveNext
    Wend
End If



'*************************************************************************

If IsNull(rs_RREQUISITN("location")) = False Then
cbo_location.Text = Trim(rs_RREQUISITN("location")) '**Add 5/7/2001'txt_location.Text = Trim(rs_RREQUISITN("location"))'**comm 5/7/2001
Else
cbo_location.Text = "" 'txt_location.Text = ""'*****changed 5/7/2001
End If

If Trim(rs_RREQUISITN("REQNSTAT")) = "M" Then
Dim MSG_MEMO
 MSG_MEMO = MsgBox("A MEMO has  ALREADY been issued for this REQUISITION.", vbInformation)
 MEMOGEN_VAR = 1
End If

If new_view_flg = 1 Then
    frm_entry.Enabled = True
    msf_calendar.Enabled = True
    cbo_apttm.Visible = True
    lbl_apttm.Visible = False
    cbo_apttm.Enabled = True
    txt_instructions.Enabled = True
    cbo_location.Enabled = True '****txt_location.Enabled = True'*changed 5/7/2001
    End If

If rs_RREQUISITN.RowCount <> 0 Then
  lbl_proccode = rs_RREQUISITN("ser_code")
End If

rq_procname(0) = lbl_proccode.Caption
If flg_procname = False Then
    Set rs_procname = rq_procname.OpenResultset(1, 1)
    flg_procname = True
Else
    rs_procname.MoreResults
    rs_procname.Requery
End If

If rs_procname.RowCount <> 0 Then
lbl_procname = rs_procname("ser_det")
End If

'******if instructions exit in RAPNTINSTR table, that record is fetched.****
'******else instructions from INST_det table are fetched.*****
If new_view_flg = 1 Then
    rq_appinstr(0) = req_chglen(cbo_reqno.Text)
Else
    rq_appinstr(0) = MSK_REQNUM
End If

If flg_appinstr = False Then
    Set rs_appinstr = rq_appinstr.OpenResultset(1, 1)
    flg_appinstr = True
Else
    rs_appinstr.MoreResults
    rs_appinstr.Requery
End If

If rs_appinstr.RowCount > 0 Then '***** instructions exist in RAPNTINSTR table
    txt_instructions.Text = rs_appinstr("INST")
Else            '**************instructions from INST_det table are fetched
    rq_serinstrct(0) = Trim(lbl_proccode.Caption)
    rq_serinstrct(1) = Trim(lbl_Modalitycode.Caption)
    If flg_serinstrct = False Then
        Set rs_serinstrct = rq_serinstrct.OpenResultset(1, 1)
        flg_serinstrct = True
    Else
        rs_serinstrct.MoreResults
        rs_serinstrct.Requery
    End If

    If rs_serinstrct.RowCount > 0 Then


         selrislnklf_qr(0) = rs_serinstrct("instrs")
         selrislnklf_rs.MoreResults
         selrislnklf_rs.Requery
         'MsgBox selrislnklf_rs.RowCount
         If selrislnklf_rs.RowCount > 0 Then
            txt_instructions.Text = ""
            selrislnklf_rs.MoveFirst
            While Not selrislnklf_rs.EOF
            txt_instructions.Text = Trim(txt_instructions.Text) & selrislnklf_rs("instext") & Chr(13)
            selrislnklf_rs.MoveNext
            Wend
         Else
             MsgBox "There are NO INSTRUCTIONS for this Procedure."
         End If
    End If

'''    If rs_serinstrct.RowCount > 0 Then
'''        txt_instructions.Text = ""
'''        While rs_serinstrct.EOF = False
'''            txt_instructions.Text = txt_instructions.Text & " / " & rs_serinstrct(0)
'''            rs_serinstrct.MoveNext
'''        Wend
'''        var_txtinstruct = txt_instructions.Text
'''    Else
'''        txt_instructions.Text = ""
'''    End If
End If

    If new_view_flg = 2 Then
        msk_caseno.Text = rs_RREQUISITN("caseno")
       ' retrieve
        'msk_REQNNO.Enabled = False'**comm 29052001
        cbo_reqno.Enabled = True '**add 29052001
        msk_caseno.Enabled = False
        cbo_deptname.Enabled = False '*******5/7/2001
        If cmd_cancel.Enabled = True Then
        cmd_cancel.SetFocus
        ElseIf cmd_view.Enabled = True Then
        cmd_view.SetFocus
        End If
    End If
  Screen.MousePointer = 1
  lbl_msg.Caption = ""
  If MEMOGEN_VAR = 1 Then
    lbl_msg.Caption = "MEMO ALREADY ISSUED FOR THIS REQUISITION!"
  End If
Exit Sub
errlbl:
MsgBox Error(Err)
Resume Next
Exit Sub
MSG:
MsgBox Error(Err)
End Sub


Private Sub cbo_reqno_LostFocus()
If cbo_location.Enabled = True Then  '*****txt changed to cbo location 5/7/2001
cbo_location.SetFocus
End If
End Sub

Private Sub cmd_cancel_Click()
cmdbtn_vis ("CA")
clear_fields
frm_key.Enabled = False
msk_caseno.Enabled = False
cbo_deptname.Enabled = False '***add 5/7/2001
'msk_REQNNO.Enabled = False'**comm 29052001
frm_disp.Enabled = False
'atk
'frm_entry.Enabled = False
cbo_reqno.Enabled = False
'msk_REQNNO.Enabled = False'***comm 29052001
End Sub

Private Sub cmd_commit_Click()
cmdbtn_vis ("CO")
i = MsgBox("The RECORD is successfully COMMITTED.", vbInformation + vbOKOnly)
End Sub

Private Sub cmd_delete_Click()
Dim i As Integer
On Error GoTo msg_delerr:

rdb.BeginTrans

reqnextstg_chk
'MsgBox "should the details be removed from RAPNTINSTR table or not?"
If reqnextstg_flg = 0 Then
    cmdbtn_vis ("D")
    i = MsgBox("Do you want to DELETE the RECORD?", vbQuestion + vbYesNo + vbDefaultButton2)
    If i = vbYes Then
        cmdbtn_vis ("YES")
     '********updating RREQUISITN table********************
        rq_RREQUISITNApptUpdt_Qr(0) = Null
        rq_RREQUISITNApptUpdt_Qr(1) = Null
        rq_RREQUISITNApptUpdt_Qr(2) = Null
        rq_RREQUISITNApptUpdt_Qr(3) = Null
        rq_RREQUISITNApptUpdt_Qr(4) = MSK_REQNUM   'tmh_ris_commonroutines.req_chglen(cbo_reqno.Text)
        rq_RREQUISITNApptUpdt_Qr.Execute 'UPDATE RREQUISITN
        '*******************************
        '******deleting record from RAPNTINSTR table***
        rq_AppInstrDel_Qr(0) = MSK_REQNUM
        rq_AppInstrDel_Qr.Execute
        '********************************
         delta102_qr(0) = msk_caseno.FormattedText
         delta102_qr(1) = lbl_aptdt.Caption
         delta102_qr(2) = lbl_proccode.Caption
         delta102_qr.Execute



      '******************************
        MsgBox "The RECORD is DELETED.", vbInformation
        clear_fields
    Else
        cmdbtn_vis ("NO")
        MsgBox "The RECORD is NOT DELETED.", vbInformation
    End If
    '*****************************

    '***********if memo issued but appt to be deleted******
 ElseIf reqnextstg_flg = 2 Then
    cmdbtn_vis ("D")
    i = MsgBox("DO YOU WANT TO DELETE THE APPOINTMENT" & Chr(13) & "EVEN IF THE MEMO IS GENERATED ?", vbQuestion + vbYesNo + vbDefaultButton2)
    If i = vbYes Then
        cmdbtn_vis ("YES")
     '********updating RREQUISITN table**ONLY APPT DATE AND TIME BEING MADE NULL******************
        rq_RREQUISITNApptUpdt_Qr(0) = Null
        rq_RREQUISITNApptUpdt_Qr(1) = Null
        rq_RREQUISITNApptUpdt_Qr(2) = Trim(cbo_location.Text)
        rq_RREQUISITNApptUpdt_Qr(3) = "M"
        rq_RREQUISITNApptUpdt_Qr(4) = MSK_REQNUM   'tmh_ris_commonroutines.req_chglen(cbo_reqno.Text)
        rq_RREQUISITNApptUpdt_Qr.Execute 'UPDATE RREQUISITN
        '*******************************
        '******deleting record from RAPNTINSTR table***
'        rq_AppInstrDel_Qr(0) = MSK_REQNUM
'        rq_AppInstrDel_Qr.Execute
        '********************************
        'deleting ta102
         delta102_qr(0) = msk_caseno.FormattedText
         delta102_qr(1) = lbl_aptdt.Caption
         delta102_qr(2) = cbo_apttm.Text
         delta102_qr.Execute

        'deleting ta102
      '******************************
        MsgBox "The RECORD is DELETED.", vbInformation
        clear_fields
    Else
        cmdbtn_vis ("NO")
        MsgBox "The RECORD is NOT DELETED.", vbInformation
    End If



    '**************
End If


rdb.CommitTrans
Exit Sub
msg_delerr:
    rdb.RollbackTrans
    MsgBox Error(Err)

End Sub

Private Sub CMD_DIALY_Click()
'If Len(Trim(tmh_ris_appointment.lbl_aptdt.Caption)) <> 0 Then
  ' tmh_ris_apptslot21.Show
  ' tmh_ris_apptslot23.show
'Else
' MsgBox "Appointment Date Not Selected"
'End If
End Sub

Private Sub cmd_exit_Click()
On Error GoTo MSG:
Me.MousePointer = 0
Me.Hide
DIS_MAIN.show
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmd_modify_Click()
cmdbtn_vis ("M")
End Sub

Private Sub cmd_new_Click()
On Error GoTo MSG:
cmdbtn_vis ("N")
new_view_flg = 1
clear_fields
'msk_REQNNO.Visible = False'**comm 29052001
cbo_reqno.Visible = True
frm_key.Enabled = True

cbo_apttm.Visible = True
lbl_apttm.Visible = False

'*****added , comm 5/7/2001*****
If Len(Trim(cbo_deptname.Text)) = 0 Then
    Dim ind_cbodept As Integer
    For ind_cbodept = 0 To (cbo_deptname.ListCount - 1)
        If Left(Trim(cbo_deptname.List(ind_cbodept)), 14) = "RADIODIAGNOSIS" Then
            cbo_deptname.ListIndex = ind_cbodept
            'cbo_deptname.Text = cbo_deptname.List(IND_CBODEPT - 1)
        End If
    Next ind_cbodept
End If

var_olddeptname = cbo_deptname.Text

'lbl_msg.Caption = "select Department for appointment is to be given"
cbo_deptname.Enabled = True
'cbo_deptname.SetFocus
msk_caseno.Enabled = True

lbl_msg.Caption = "enter case number whose appointment is to be given"
'msk_caseno.SetFocus
No_Chargeflag = False
'*************
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmd_print_Click()

cmdbtn_vis ("P")
Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim str As String
Dim INST As String
Dim i As Integer
Dim beforst As String, afterst As String
Dim K As Long
On Error GoTo prtmsg:
lbl_msg.Caption = "PRINTING..."
DoEvents

    Cr_appointment.SelectionFormula = "{RREQUISITN.REQNNO} = '" & REQNUM & "'"
    Cr_appointment.Connect = "UID=PATOOLS;PWD=PATOOLS;Database = EC_TMH"

      'CR2.Action = 1
'      CR2.ReportFileName = "C:\PB121298\adm_doc1.rpt"
'      CR2.Destination = crptToPrinter
'      CR2.PrintReport

   If lbl_wrdbdno.Visible = True Then
    Cr_appointment.Formulas(0) = "ward_bed = '" & lbl_wrdbdno.Caption & "/" & lbl_bedno.Caption & "'"
    Cr_appointment.Formulas(1) = "wardbed_lbl ='" & "WARD NAME / BED NO." & "'"
   Else
    Cr_appointment.Formulas(0) = "ward_bed =''"   '& "'"
    Cr_appointment.Formulas(1) = "wardbed_lbl =''"   '& "'"
   End If

    If cbo_apttm.Visible = True Then
      INST = Left(cbo_apttm.Text, 10)
        Cr_appointment.Formulas(2) = "apttm = '" & INST & "'"
    ElseIf lbl_apttm.Visible = True Then
        INST = Left(lbl_apttm.Caption, 11)
        Cr_appointment.Formulas(2) = "apttm = '" & INST & "'"
    End If

   If lbl_ipopsts.Caption = "Out-patient" Then
    Cr_appointment.Formulas(3) = "patipopstat = '" & "OP" & "'"
   Else
    Cr_appointment.Formulas(3) = "patipopstat = '" & "IP" & "'"
   End If
    Cr_appointment.Formulas(4) = "REQNO = '" & req_chglen(REQNUM) & "'"
    If Trim(lbl_Modalitycode.Caption) = "CT" Or Trim(lbl_Modalitycode.Caption) = "CR" Or Trim(lbl_Modalitycode.Caption) = "MM" Or Trim(lbl_Modalitycode.Caption) = "MR" Or Trim(lbl_Modalitycode.Caption) = "US" Then
         Cr_appointment.Formulas(5) = "deptname = '" & "Department of Radiodiagnosis" & "'"
    End If
    If Trim(lbl_Modalitycode.Caption) = "2D" Or Trim(lbl_Modalitycode.Caption) = "ST" Or Trim(lbl_Modalitycode.Caption) = "PF" Then
         Cr_appointment.Formulas(5) = "deptname = '" & "Department of General Medicine " & "'"
    End If
    If Trim(lbl_Modalitycode.Caption) = "ED" Then
         Cr_appointment.Formulas(5) = "deptname = '" & "Department of ENDOSCOPY " & "'"
    End If
 '   cr_appointment.Formulas(6) = "instruction = '" & Trim(txt_instructions.Text) & "'"

    str1 = ipadd
     str2 = Format(DATE11, "dd/mm/yyyy")
     STR3 = time1
    str1 = str1 & Space(5) & str2 & Space(5) & STR3
    str1 = Space(3) & str1
    Cr_appointment.Formulas(6) = "us_id = '" & str1 & "'"
    If Len(Trim(msk_lmpdate.FormattedText)) > 4 Then
         Cr_appointment.Formulas(7) = "lmp = '" & "LMP " & msk_lmpdate.FormattedText & "'"
    Else
         Cr_appointment.Formulas(7) = "lmp = '" & " " & "'"

    End If

    str = Format(lbl_aptdt.Caption, "dddd")
    Cr_appointment.Formulas(8) = "week = '" & str & "'"



    INST = txt_instructions.Text



    'cr_appointment.Formulas(9) = "instructions = '" & INST & "'"


    Cr_appointment.ReportFileName = "C:\DISREPORTSEC\crpt_appt1.rpt"
    Cr_appointment.WindowState = crptMaximized
    Cr_appointment.Destination = crptToWindow
    Cr_appointment.Action = 1
lbl_msg.Caption = ""

Exit Sub
prtmsg: MsgBox Error(Err)
lbl_msg.Caption = ""

End Sub

Private Sub cmd_query_Click()
On Error GoTo MSG:
If Trim(lbl_aptdt) = "" Then
MsgBox "Please SELECT  the DATE."
Exit Sub
Else
QRYFLAG = True
tmh_ris_apptqry.show
tmh_ris_apptqry.msk_aptdate.Text = lbl_aptdt
Dim ind_cbodept As Integer
For ind_cbodept = 0 To (tmh_ris_apptqry.cbo_deptname.ListCount - 1)
        If Trim(tmh_ris_apptqry.cbo_deptname.List(ind_cbodept)) = "RADIODIAGNOSIS$LZZ" Then
            tmh_ris_apptqry.cbo_deptname.ListIndex = ind_cbodept
        End If
    Next ind_cbodept

tmh_ris_apptqry.cbo_modality.SetFocus
End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmd_save_Click()
On Error GoTo MSG:
Dim j As Integer
msk_systime.Text = Format(Time, "hh:mm:ss")


If Len(Trim(msk_caseno.Text)) = 0 Then
j = MsgBox("The CASE NUMBER field is EMPTY.", vbCritical)
Exit Sub
End If

If Len(Trim(cbo_reqno.Text)) = 0 Then
j = MsgBox("The REQUISITION NUMBER field is EMPTY.", vbCritical)
cbo_reqno.SetFocus
Exit Sub
End If


If Len(Trim(lbl_aptdt.Caption)) = 0 Then
j = MsgBox("The APPOINTMENT DATE field is EMPTY.", vbCritical)
Exit Sub
End If

If Len(Trim(cbo_apttm.Text)) = 0 Then
j = MsgBox("The APPOINTMENT TIME field is EMPTY.", vbCritical)
Exit Sub
End If

'If IsNumeric(Trim(cbo_apttm.Text)) = False Then
'j = MsgBox("Appointment time field is entered wrong", vbCritical)
'Exit Sub
'End If

If Len(Trim(cbo_location.Text)) = 0 Then '**changed 5/7/2001** If Len(Trim(txt_location.Text)) = 0 Then
j = MsgBox("Location field is empty" & Chr(13) & "enter '-' to leave the filed blank", vbCritical)
cbo_location.SetFocus '**changed 5/7/2001  txt_location.SetFocus '*****add 29052001
Exit Sub
End If

assign
'i = MsgBox("Record successfully Added / Updated", vbInformation + vbOKOnly)
cmdbtn_vis ("S")
frm_key.Enabled = False
'atk
frm_entry.Enabled = False
Exit Sub
MSG:
MsgBox Error(Err)
End Sub


Private Sub Cmd_help_Click()
'appt_sts.Show
'tmh_ris_apptslot1.Show
'tmh_ris_apptslot2.show
End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG:
cmdbtn_vis ("V")
new_view_flg = 2
clear_fields
'cbo_reqno.Visible = False'**comm 29052001
'msk_REQNNO.Visible = True'***comm 29052001
cbo_reqno.Visible = True '**add 29052001

frm_key.Enabled = True
cbo_reqno.Enabled = True '***add 29052001'msk_REQNNO.Enabled = True'**comm 29052001


'frm_disp.Enabled = True
'frm_key.Enabled = False
frm_entry.Enabled = True
msf_calendar.Enabled = True
cbo_apttm.Enabled = False
cbo_apttm.Visible = False
lbl_apttm.Visible = True
txt_instructions.Enabled = False
cbo_location.Enabled = False '**txt_location.Enabled = False'**chnaged 5/7/2001

'*****added , comm 5/7/2001*****
If Len(Trim(cbo_deptname.Text)) = 0 Then
    Dim ind_cbodept As Integer
    For ind_cbodept = 0 To (cbo_deptname.ListCount - 1)
        If Left(Trim(cbo_deptname.List(ind_cbodept)), 14) = "RADIODIAGNOSIS" Then
            cbo_deptname.ListIndex = ind_cbodept
            'cbo_deptname.Text = cbo_deptname.List(IND_CBODEPT - 1)
        End If
    Next ind_cbodept
End If

var_olddeptname = cbo_deptname.List(ind_cbodept - 1)
'lbl_msg.Caption = "select Department for appointment is to be given"
cbo_deptname.Enabled = True
msk_caseno.Enabled = True
lbl_msg.Caption = "Enter case number for whom the appointment is to be given"
'msk_caseno.SetFocus

'cbo_deptname.SetFocus
'*5/7/2001 add,; comm'
Exit Sub
MSG:
MsgBox Error(Err)
End Sub


Private Sub Command1_Click()
     VALID_DT
End Sub

Private Sub Form_Activate()
On Error GoTo MSG:
'If appslt_viewed = 1 Then
'    If new_view_flg = 1 Then
'    cmdbtn_vis ("N")
'    frm_entry.Enabled = True
'    End If
'Else
'    cmdbtn_vis ("I")
'    frm_key.Enabled = False
'    'atk
'    'frm_entry.Enabled = False
'    cmd_new.SetFocus
'
'End If

'***** comm 29052001 from here
'If new_view_flg = 1 Then
'cbo_reqno.Visible = True
'msk_REQNNO.Visible = False
'Else
'cbo_reqno.Visible = False
'msk_REQNNO.Visible = True
'End If
'******comm 29052001 till here


If new_view_flg <> 2 Then
cbo_apttm.Visible = True
lbl_apttm.Visible = False
Else
lbl_apttm.Visible = True
cbo_apttm.Visible = False
End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub Form_Load()
On Error GoTo MSG:
'Set RS = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107") '***2/2/99
datetime
'DATE11 = RS("dt1") '**add 24/11/2k
'time1 = Format(RS("TM1"), "hh:mm:ss") '***2/2/99
msk_sysdate.Text = DATE11
msk_systime.Text = time1
new_view_flg = 0
var_aptnotgiven = 0
var_nopendingreq = 0
'******from req fh
Dim str As String
Screen.MousePointer = 0
No_Chargeflag = False
If APPOINTMENTFLAG = False Then
    APPOINTMENTFLAG = True
'str = "select REQNNO,REQNSTAT from RREQUISITN where caseno = ? and (isnull(REQNSTAT) or trim(REQNSTAT) = '')"'**isnull(x) works for msaccess; x is null for as/400
'str = "select REQNNO,REQNSTAT from RREQUISITN where caseno = ? and (REQNSTAT IS NULL or trim(REQNSTAT) = '')"'**comm 5/7/2001 gives for all depatments
'str = "select REQNNO,REQNSTAT from RREQUISITN where caseno = ? and dept_code = ?  and (REQNSTAT IS NULL or trim(REQNSTAT) = '')" '**add 5/7/2001 for selected department
'31/10/2001-Parag
'This is the previous str statement
'STR = "select REQNNO,REQNSTAT,REQNDT from RREQUISITN where caseno = ? and dept_code = ?  and ((REQNSTAT IS NULL or trim(REQNSTAT) = '') OR (REQNSTAT = 'M' AND APNTDT IS NULL)) order by reqndt desc " '**add 4/9/2001 for DELETING APPT EVEN IF MEMO IS GENERATED
'This is the STR statement in which all requisitions where REQNSTAT <> 'E' or REQNSTAT <> 'F' or REQNSTAT <> 'T' REQNSTAT <> 'A' that is
'after the examination is done will not be considered.
str = "select REQNNO,REQNSTAT,REQNDT from rislib.RREQUISITN where caseno = ? and dept_code = ?  and ( (REQNSTAT IS NULL or trim(REQNSTAT) = '' or trim(REQNSTAT) = 'M' or trim(REQNSTAT) = 'P' ))AND APNTDT IS NULL  order by reqndt desc "

Set rq_reqpend = rdb.CreateQuery("rq_reqpend_apptstmt", str)

'str = "select REQNNO,REQNSTAT from RREQUISITN where caseno = ? and (not isnull(REQNSTAT) or trim(REQNSTAT) <> '')"'**msaccess
'str = "select REQNNO,REQNSTAT from RREQUISITN where caseno = ? and (not isnull(REQNSTAT) or trim(REQNSTAT) <> '') and APNTDT >= ?"''**for future appts only; msaccess
'str = "select REQNNO,REQNSTAT from RREQUISITN where caseno = ? and  (REQNSTAT IS NOT NULL or trim(REQNSTAT) <> '') and APNTDT >= ?" '**for future appts only; as/400'**all departments
str = "select REQNNO,REQNSTAT from rislib.RREQUISITN where caseno = ? and dept_code = ? and  (REQNSTAT IS NOT NULL or trim(REQNSTAT) <> '') and APNTDT >= ?" '**for future appts only; as/400'****for selected dept

Set rq_apptgiven = rdb.CreateQuery("rq_apptgiven_apptstmt", str)

'****5/7/2001 added
'str = "select dept_code,dept_name from PATOOLSLIB.ma106 where dept_code in(select distinct dept_code from RMODALITY)"'****department for which modalities are present
str = "SELECT * FROM PATOOLSLIB.MA106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ' OR DEPT_CODE='IZZ' ) "
'STR = "select dept_code,dept_name from PATOOLSLIB.ma106" '*********all departments.
Set rq_dept = rdb.CreateQuery("rq_dept_apptstmt", str)
'****


'str = "select dept_name from PATOOLSLIB.ma106"
'Set rq_dept = rdb.CreateQuery("rq_dept_apptstmt", str)
'
'str = "select dept_code from PATOOLSLIB.ma106 where dept_name = ?"
'Set rq_deptcode = rdb.CreateQuery("rq_deptcode_apptstmt", str)

str = "select tag,first_name,mid_name,last_name,birth_date,sex,pat_reg_ty,pat_catg, ud_name from PATOOLSLIB.ta101 where  PATOOLSLIB.ta101.case_no = ?"  '(select caseno from ex_hms where caseno = ?)"
'str = "select tag,first_name,mid_name,last_name,birth_date,sex,pat_catg, ref_by_doc from PATOOLSLIB.ta101, PATOOLSLIB.ta1011 where PATOOLSLIB.ta101.case_no = PATOOLSLIB.ta1011.case_no and PATOOLSLIB.ta101.case_no = ?"  '(select caseno from ex_hms where caseno = ?)"
'str = "select tag,first_name,mid_name,last_name,birth_date,sex,pat_catg, ref_by_doc, status from PATOOLSLIB.ta101, PATOOLSLIB.ta1011, PATOOLSLIB.ma104 where PATOOLSLIB.ta101.case_no = PATOOLSLIB.ta1011.case_no and PATOOLSLIB.ta101.case_no = PATOOLSLIB.ma104.case_no  and PATOOLSLIB.ta101.case_no = ?"  '(select caseno from ex_hms where caseno = ?)"
Set rq_patdet = rdb.CreateQuery("rq_patdet_apptstmt", str)

str = "select status, ward_name, bed_no from PATOOLSLIB.ma104 where PATOOLSLIB.ma104.case_no  = ?"
Set rq_patipop = rdb.CreateQuery("rq_patipop_apptstmt", str)

'str = "select distinct MODLITYDES from RMODALITY"
'Set rq_modality = rdb.CreateQuery("rq_modality_apptstmt", str)

str = "select MODLITYDES from rislib.RMODALITY where MODALITYCD = ?"
Set rq_moddes = rdb.CreateQuery("rq_moddes_apptstmt", str)

'str = "select ud_name from PATOOLSLIB.ma1031"
'Set rq_refdoc = rdb.CreateQuery("rq_refdoc_apptstmt", str)

'str = "select ser_det from PATOOLSLIB.ma101"
'Set rq_proc = rdb.CreateQuery("rq_proc_apptstmt", str)

str = "select ser_det from PATOOLSLIB.ma101 where ser_code = ?"
Set rq_procname = rdb.CreateQuery("rq_procname_apptstmt", str)

'str = "select prvdiagdes from PATOOLSLIB.ta1012 where case_no = ?"
'Set rq_provdiag = rdb.CreateQuery("rq_provdiag_apptstmt", str)

'******add 5/7/2001 for location combo box*****
str = "select location from rislib.rlocation where dept_code = ? and modalitycd =?"
Set rq_applocation = rdb.CreateQuery("rq_applocation_apptstmt", str)
'************

str = "INSERT INTO rislib.RREQUISITN (REQNNO, REQNDT,REQNTM,CASENO,DEPT_CODE,MODALITYCD,PATNSTAT,PAT_CATG,SER_CODE,REFBYDOC,EXAMINATN,PREVEXAM) VALUES " _
        & "(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
Set RREQUISITNInsrt_Qr = rdb.CreateQuery("RREQUISITNInsrt_Qrappt", str)

Set Insrt_Stmt = rdb.CreateQuery("INSTMT", "INSERT INTO PATOOLSLIB.TA102 (CASE_NO,APPT_DATE,APPT_TIME,INV_TYPE,INV1,INV2,INV3,INV4,INV5,REF_DOC,BY_DOC,FEE_DATE,IS_DATE) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")
'Set Insrt_Stmt = rdb.CreateQuery("INSTMT", "INSERT INTO PATOOLSLIB.TA102 (CASE_NO,APPT_DATE,APPT_TIME,INV_TYPE,INV1,INV2,INV3,INV4,INV5,REF_DOC,BY_DOC,FEE_DATE,IS_DATE) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")
'str = "select * from RREQUISITN where REQNNO = ?"
str = "select REQNNO,caseno,ser_code,REQNSTAT,APNTDT,APNTTM,location,dept_code,refbydoc from rislib.RREQUISITN where REQNNO = ?"
Set rq_RREQUISITN = rdb.CreateQuery("rq_RREQUISITN_apptstmt", str)



'str = "update ma104 set status = ?, case_no = ?, bd_free_dt = ? where (trim(ward_name) = ? and trim(BED_INCH) = ? and trim(bed_no) = ?)"
'Set Ma104DelU_Qr = rdb.CreateQuery("BDMa104Del_Q3", str)


'******till here from req frm

'str = "select INSTRS from RINVESTIGN where ser_code = ?"


str = "select instrs,modalitycd,ser_code from rislib.RINVESTIGN where ser_code = ? and MODALITYCD = ?"
Set rq_serinstrct = rdb.CreateQuery("rq_serinstrct_apptstmt", str)

str = "select INST, REQNNO from rislib.RAPNTINSTR where REQNNO = ?"
Set rq_appinstr = rdb.CreateQuery("rq_appinstr_apptstmt", str)

'str = "update RREQUISITN set APNTDT = ?, APNTTM = ?, REQNSTAT = ? where REQNNO = ?"
'Set rq_RREQUISITNApptInsrt_Qr = rdb.CreateQuery("rq_RREQUISITNApptInsrt_apptstmt", str)
'
'str = "update RREQUISITN set APNTDT = ?, APNTTM = ?, REQNSTAT = ? where REQNNO = ?"
'Set rq_RREQUISITNDelete_Qr = rdb.CreateQuery("rq_RREQUISITNDel_apptstmt", str)

str = "update RISLIB.RREQUISITN set APNTDT = ?, APNTTM = ?, location = ?, REQNSTAT = ? where REQNNO = ?"
Set rq_RREQUISITNApptUpdt_Qr = rdb.CreateQuery("rq_RREQUISITNApptUpdt_apptstmt", str)

str = "insert into RISLIB.RAPNTINSTR values(?,?,?)"
Set rq_AppInstrInsrt_Qr = rdb.CreateQuery("rq_AppInstrInsrt_apptstmt1", str)

str = "delete from rislib.RAPNTINSTR where REQNNO = ?"
Set rq_AppInstrDel_Qr = rdb.CreateQuery("rq_AppInstrDel_apptstmt", str)

str = "select hday from PATOOLSLIB.ma1121"
Set rq_calfill = rdb.CreateQuery("rq_calfill_apptstmt", str)

str = "select NAFROMDT,NATODT,NAREASON from rislib.RSERNONAVL where MODALITYCD = ? and dept_code = ? and NATODT>= ?"
'str = "select NAFROMDT,NATODT from RSERNONAVL where MODALITYCD = ? and dept_code = ? and NATODT>= ?"
Set rq_sernonavail = rdb.CreateQuery("rq_sernonavail_apptstmt", str)


str = "select lmp_date from rislib.CLINHIST where CASENO = ?" '******15/09/2001
Set rq_clinhistlmpdt = rdb.CreateQuery("rq_clinhistlmpdt_apptstmt", str)


str = "select  caseno, APNTDT,APNTTM,MODALITYCD  from rislib.RREQUISITN where APNTDT = ? and CASENO = ?  and dept_code = ?"
Set rq_caseapptgiven = rdb.CreateQuery("RQ_CASEAPPTGIVEN_STMT", str)

str = "select grpno,insno,instext from rislib.rislnklf where grpno= ?"
Set selrislnklf_qr = rdb.CreateQuery("selrislnklf_qr1", str)
selrislnklf_qr(0) = 0
Set selrislnklf_rs = selrislnklf_qr.OpenResultset(1, 1)


'STR = "SELECT * FROM RISLIB.RISAUTH WHERE USRGROUP=? AND SCRN_DESC=?"
'Set RISAUTH_QR = rdb.CreateQuery("SELRISAUTH21_QR1", STR)
'RISAUTH_QR(0) = Trim(USRGRP)
'RISAUTH_QR(1) = "APPOINTMENT"
'Set RISAUTH_RS = RISAUTH_QR.OpenResultset(2, 3)

'If RISAUTH_RS.RowCount = 0 Then
'MsgBox "No Button Authorisation for this User Group & Screen"
'End
'End If


'If appslt_viewed = 1 Then
'    If new_view_flg = 1 Then
    cmdbtn_vis ("N")
    frm_entry.Enabled = True
'    End If
'Else
    cmdbtn_vis ("I")
    frm_key.Enabled = False
    'atk
    'frm_entry.Enabled = False
    'cmd_new.SetFocus

'End If




setflagsfalse
reqnextstg_flg = 0
MEMOGEN_VAR = 0 '*****4/9/2001 ADD

CALENDAR_FILL
'**********
'******add 5/7/2001
If flg_dept = False Then
    Set rs_dept = rq_dept.OpenResultset(1, 1)
    flg_dept = True
Else
    rs_dept.MoreResults
    rs_dept.Requery
End If

cbo_deptname.clear
If rs_dept.RowCount > 0 Then
    While rs_dept.EOF = False
        cbo_deptname.AddItem Trim(rs_dept("dept_name")) & "$" & rs_dept("dept_code")
        rs_dept.MoveNext
    Wend
End If

str = "select c_ch,nc_ch from PATOOLSLIB.ma101 where ser_code=?"
Set SELMA101_QR = rdb.CreateQuery("selma101_qr1", str)
SELMA101_QR(0) = "0000"
Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 1)

str = "delete from PATOOLSLIB.ta102 where case_no=? and appt_date=? and inv1=?"
Set delta102_qr = rdb.CreateQuery("delta102_qr1", str)

'**********added 5/7/2001 till here

str = "SELECT * FROM RISLIB.RREQUISITN WHERE reqnDT=? AND MODALITYCD=?"
Set selreq_qr = rdb.CreateQuery("SELREQ_QR1", str)
selreq_qr(0) = "2001-01-01"
selreq_qr(1) = "XX"
Set selreq_rs = selreq_qr.OpenResultset(1, 3)

':::::::::::::::
End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Sub CALENDAR_FILL()
On Error GoTo MSG:
cbo_month.AddItem "JANUARY"
cbo_month.AddItem "FEBRUARY"
cbo_month.AddItem "MARCH"
cbo_month.AddItem "APRIL"
cbo_month.AddItem "MAY"
cbo_month.AddItem "JUNE"
cbo_month.AddItem "JULY"
cbo_month.AddItem "AUGUST"
cbo_month.AddItem "SEPTEMBER"
cbo_month.AddItem "OCTOBER"
cbo_month.AddItem "NOVEMBER"
cbo_month.AddItem "DECEMBER"

i = 2000
While cbo_year.ListCount <= 200
cbo_year.AddItem i
i = i + 1
Wend


cbo_month.Text = UCase(Format(DATE11, "MMMM"))
cbo_year.Text = Format(DATE11, "YYYY")
'lbl_monthyear.Caption = Format(Date, "MMMM,YYYY") '**comm 4/7/2001
lbl_monthyear.Caption = "" '**add 4/7/2001
m = Format(DATE11, "mm")
Y = Format(DATE11, "yyyy")
'Main


If flg_calfill = False Then
    Set rs_calfill = rq_calfill.OpenResultset(1, 1)
    flg_calfill = True
Else
    rs_calfill.MoreResults
    rs_calfill.Requery
End If

'Set rs_calfill = rdb.OpenResultset("select hday from PATOOLSLIB.ma1121", 1, 3)
totholidays = rs_calfill.RowCount
ReDim hlist(rs_calfill.RowCount) As String
i = 0
While rs_calfill.EOF = False
hlist(i) = Format(rs_calfill(0), "yyyy/mm/dd")
i = i + 1
rs_calfill.MoveNext
Wend
f1
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Sub fillcal_sernonavail()
On Error GoTo MSG:
'Dim totnonavldays As Integer
Screen.MousePointer = 11
'commented for ris
rq_sernonavail(0) = lbl_Modalitycode.Caption
rq_sernonavail(1) = Mid(cbo_reqno.Text, 1, 3)
rq_sernonavail(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY")) 'CDate(Format(Date11, "yyyy/mm/dd"))
If flg_sernonavail = False Then
    Set rs_sernonavail = rq_sernonavail.OpenResultset(1, 1)
    flg_sernonavail = True
Else
    rs_sernonavail.MoreResults
    rs_sernonavail.Requery
End If
'commented for ris
'str = "select NAFROMDT,NATODT from RSERNONAVL where MODALITYCD ='" & lbl_modalitycode.Caption & "'  and dept_code ='" & Mid(Cbo_reqno.Text, 1, 3) & "' and NATODT >= " & CDate(Format(Date, "yyyy/mm/dd"))
'str = "select NAFROMDT,NATODT from RSERNONAVL where MODALITYCD = ? and dept_code = ? and NATODT >= ?"

'Set rs_sernonavail = rdb.OpenResultset(str, 1, 3)


'Set rs_sernonavail = rdb.OpenResultset("select NAFROMDT,NATODT from RSERNONAVL where MODALITYCD = '" & lbl_modalitycode.Caption & "' and dept_code='" & Mid(cbo_reqno.Text, 1, 3) & "' and todate>=" & CDate(Format(Date, "yyyy/mm/dd")) & "", 1, 3)
NVLROWS = rs_sernonavail.RowCount
ReDim fromdt(rs_sernonavail.RowCount) As String
ReDim todt(rs_sernonavail.RowCount) As String
i = 0
While rs_sernonavail.EOF = False
    todt(i) = Format(rs_sernonavail("NATODT"), "dd/mm/yyyy")
    fromdt(i) = Format(rs_sernonavail("NAFROMDT"), "dd/mm/yyyy")
    i = i + 1
    rs_sernonavail.MoveNext
Wend

If Left(rs_sernonavail("NAREASON"), 2) = "P," Then
MsgBox "Please give only Limited Appointments.Please Check with the Head Of Department."
End If
'    totnonavldays = DateDiff("d", rs_sernonavail(0), rs_sernonavail(1))
'    ReDim srvnonavail_list(totnonavldays) As String
'
'    For i = 0 To totnonavldays
'    srvnonavail_list(i) = Format(DateAdd("d", i, rs_sernonavail(0)), "yyyy/mm/dd")
'   ' MsgBox srvnonavail_list(i)
'    Next
'   ' Exit Sub
'   f2
'Else
'    f1
'End If
Screen.MousePointer = 0
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Sub clear_fields()

msk_caseno.Mask = ""
msk_caseno.Text = ""
' msk_caseno.Mask = ">??/#####"

'****comm 29052001 from here
'msk_REQNNO.Mask = ""
'msk_REQNNO.Text = ""
'msk_REQNNO.Mask = ">??A/??/##/######" '">??#/??/##/######"
'*****comm till here

cbo_deptname.ListIndex = -1 '***5/7/2001

cbo_reqno.ListIndex = -1
cbo_reqno.Text = ""

lbl_modality.Caption = ""
lbl_Modalitycode.Caption = ""
'cbo_deptname.ListIndex = -1
'lbl_deptname.Caption = ""
'lbl_deptcode.Caption = ""

lbl_name.Caption = ""
lbl_age.Caption = ""
lbl_sex.Caption = ""
lbl_patregtype.Caption = ""
lbl_catg.Caption = ""
lbl_ipopsts.Caption = ""
lbl_refby.Caption = ""
lbl_wrdbdno.Caption = ""
lbl_bedno.Caption = "" '**add 4/7/2001
lbl_procname.Caption = ""
lbl_proccode.Caption = ""

clear_entryframe
lbl_msg.Caption = ""
End Sub
Sub clear_entryframe()
'    cbo_modality.ListIndex = -1
'    lbl_modcode.Caption = ""
'    cbo_refdocser.ListIndex = -1
'    cbo_procname.ListIndex = -1
    lbl_proccode.Caption = ""
    txt_instructions.Text = ""
    'txt_location.Text = ""'**comm 5/7/2001
    cbo_location.Text = "" '**add 5/7/2001
    lbl_aptdt.Caption = ""
    'cbo_apttm.Text = ""
    lbl_apttm.Caption = ""
    cbo_apttm.ListIndex = -1

'    txt_clinchist.Text = ""
'    txt_examreq.Text = ""
'    txt_provdiag.Text = ""
'    cbo_yesno.ListIndex = -1
End Sub



Sub retrieve()
'*****8
On Error GoTo MSG:
LMPDT_VIS = 0

Dim cno_pad As String
'cno_pad = padzero(5, "C", msk_caseno.Text, 3)
msk_caseno.Text = cno_pad
CASENUM = cno_pad
var_nopendingreq = 0

 DEPTCODE_FIND

'********retieving reqnos for that case no.*******
If new_view_flg = 1 Then  '******new appt

    rq_reqpend(0) = cno_pad 'msk_caseno.Text
    rq_reqpend(1) = VAR_DCD ''**5/7/2001
    '********

    If flg_reqpend = False Then
        Set rs_reqpend = rq_reqpend.OpenResultset(1, 1)
        flg_reqpend = True
    Else
        rs_reqpend.MoreResults
        rs_reqpend.Requery
    End If
    If rs_reqpend.RowCount = 0 Then
         MsgBox "There are NO pending REQUISITIONS.", vbCritical
         lbl_msg.Caption = ""
         var_nopendingreq = 1
         Exit Sub
    Else
        cbo_reqno.clear
        rs_reqpend.MoveFirst
        cbo_reqno.Text = req_chglen(rs_reqpend(0))
        While Not rs_reqpend.EOF
            cbo_reqno.AddItem req_chglen(rs_reqpend(0))
            rs_reqpend.MoveNext
        Wend
    End If
End If


'********to view appointments already given
If new_view_flg = 2 Then  '******old appt
    rq_apptgiven(0) = cno_pad 'msk_caseno.Text
    rq_apptgiven(1) = VAR_DCD ''**5/7/2001
    rq_apptgiven(2) = CDate(GET_REVDATE(msk_sysdate.Text)) '*****29052001

    If flg_apptgiven = False Then
        Set rs_apptgiven = rq_apptgiven.OpenResultset(1, 1)
        flg_apptgiven = True
    Else
        rs_apptgiven.MoreResults
        rs_apptgiven.Requery
    End If
    If rs_apptgiven.RowCount = 0 Then
         MsgBox "There are NO future APPOINTMENTS.", vbCritical
         cbo_reqno.clear '***Add 29052001
'         msk_caseno.SelLength = 8
'         msk_caseno.SelStart = 1
'         msk_caseno.SetFocus
         var_nopendingreq = 1
         Exit Sub
    Else
        cbo_reqno.clear
        rs_apptgiven.MoveFirst
        cbo_reqno.Text = req_chglen(rs_apptgiven(0))
        While Not rs_apptgiven.EOF
            cbo_reqno.AddItem req_chglen(rs_apptgiven(0))
            rs_apptgiven.MoveNext
        Wend
    End If
End If

'*********************************


lbl_msg.Caption = "retrieving patient details...."
DoEvents
Screen.MousePointer = 11

   rq_patdet(0) = cno_pad 'msk_caseno.Text
   If flg_patdet = False Then
        Set rs_patdet = rq_patdet.OpenResultset(1, 1)
        flg_patdet = True
   Else
        rs_patdet.MoreResults
        rs_patdet.Requery
   End If

   If rs_patdet.RowCount = 0 Then
       MsgBox "The Patient is NOT REGISTERED.", vbCritical
       Screen.MousePointer = 1
       lbl_msg.Caption = ""
      ' msk_caseno.SelLength = 8
       msk_caseno.SetFocus
       Exit Sub
   End If
   lbl_name.Caption = rs_patdet("tag") & " " & rs_patdet("first_name") & " " & rs_patdet("mid_name") & " " & rs_patdet("last_name")
   lbl_sex.Caption = rs_patdet("sex")
   lbl_age.Caption = DateDiff("yyyy", rs_patdet("birth_date"), msk_sysdate.Text)
   lbl_patregtype.Caption = rs_patdet("pat_reg_ty")
   lbl_catg.Caption = " " & rs_patdet("pat_catg")
   lbl_refby.Caption = rs_patdet("ud_name")

            '******15/09/2001
        If ((Trim(lbl_sex.Caption) = "F") And (lbl_age.Caption > 13 And lbl_age.Caption < 45)) Then
            lbl_lmp.Visible = True
            msk_lmpdate.Visible = True
            LMPDT_VIS = 1

            rq_clinhistlmpdt(0) = cno_pad
            If flg_clinhistlmpdt = False Then
                Set rs_clinhistlmpdt = rq_clinhistlmpdt.OpenResultset(1, 1)
                flg_clinhistlmpdt = True
            Else
                rs_clinhistlmpdt.MoreResults
                rs_clinhistlmpdt.Requery
            End If

            If rs_clinhistlmpdt.RowCount > 0 Then
                rs_clinhistlmpdt.MoveLast
                'If IsNull(rs_clinhistlmpdt("lmp_date")) Then
                'msk_lmpdate.Text = "--"
                'Else
                If Len(rs_clinhistlmpdt("lmp_date")) <> 0 Then
                msk_lmpdate.Text = Format(rs_clinhistlmpdt("lmp_date"), "DD/MM/YYYY")
                End If
            End If



        Else
            lbl_lmp.Visible = False
            msk_lmpdate.Visible = False
            LMPDT_VIS = 0
        End If




        '*****************



   rq_patipop(0) = cno_pad  'msk_caseno.Text
   If flg_patipop = False Then
        Set rs_patipop = rq_patipop.OpenResultset(1, 1)
        flg_patipop = True
   Else
        rs_patipop.MoreResults
        rs_patipop.Requery
   End If

   If rs_patipop.RowCount = 0 Then
    Label13.Visible = False
    Label24.Visible = False
    lbl_wrdbdno.Visible = False
    lbl_bedno.Visible = False '**add 4/7/2001
    lbl_ipopsts.Caption = "Out-patient"
   Else
    Label13.Visible = True
    Label24.Visible = True
    lbl_wrdbdno.Visible = True
    lbl_bedno.Visible = True '**add 4/7/2001
    If Trim(rs_patipop("status")) = "O" Then
     lbl_ipopsts.Caption = "In-patient"   'rs_patipop("status")
    ElseIf Trim(rs_patipop("status")) = "DI" Then
     lbl_ipopsts.Caption = "In-patient" & " " & rs_patipop("status")
    ElseIf Trim(rs_patipop("status")) = "AI" Then
     lbl_ipopsts.Caption = "Out-patient" & " " & rs_patipop("status")
    End If
    'lbl_wrdbdno.Caption = rs_patipop("ward_name") & " / " & rs_patipop("bed_no")'**comm 4/7/2001
    lbl_wrdbdno.Caption = rs_patipop("ward_name") '**add 4/7/2001
    lbl_bedno.Caption = rs_patipop("bed_no") '**add 4/7/2001

   End If

  Screen.MousePointer = 0
  lbl_msg.Caption = ""

'If new_view_flg = 2 Then
'    rs_RREQUISITN(0) = cbo_reqno.Text
'    If flg_RREQUISITN = False Then
'        Set rs_RREQUISITN = rq_RREQUISITN.OpenResultset(1, 1)
'        flg_RREQUISITN = True
'    Else
'        rs_RREQUISITN.MoreResults
'        rs_RREQUISITN.Requery
'    End If
'End If
Screen.MousePointer = 0
Exit Sub

MSG:
MsgBox Error(Err)
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

Private Sub lbl_aptdt_Change()
On Error GoTo MSG:
If new_view_flg <> 2 Then
If Len(Trim(lbl_aptdt.Caption)) > 0 Then
    VALID_DT
    If calvar_holiday = 1 Then
        lbl_aptdt.Caption = ""
        msf_calendar.SetFocus
    End If

    '*******15092001*****checking if any appts given for the same case num. on the same day*****
  If Len(Trim(lbl_aptdt.Caption)) <> 0 Then
    DEPTCODE_FIND

    rq_caseapptgiven(0) = GET_REVDATE(lbl_aptdt.Caption)
    rq_caseapptgiven(1) = CASENUM
    rq_caseapptgiven(2) = VAR_DCD


    If flg_caseapptgiven = False Then
        Set rs_caseapptgiven = rq_caseapptgiven.OpenResultset(1, 1)
        flg_caseapptgiven = True
    Else
        rs_caseapptgiven.MoreResults
        rs_caseapptgiven.Requery
    End If

     Dim MSG
     Dim LST_APPTTM_MOD As String
     LST_APPTTM_MOD = ""
     If rs_caseapptgiven.RowCount > 0 Then

     While Not rs_caseapptgiven.EOF
        LST_APPTTM_MOD = LST_APPTTM_MOD & Chr(13) & rs_caseapptgiven("APNTTM") & "      " & rs_caseapptgiven("MODALITYCD")
        rs_caseapptgiven.MoveNext
     Wend


     MSG = MsgBox("APPOINTMENT ALREADY GIVEN FOR THIS CASE NUMBER" & Chr(13) & "ON THE SAMEDAY IN THIS DEPARTMENT" & Chr(13) _
                    & Chr(13) & "TIME              MODALITY" & Chr(13) & "----------------------------------------" & LST_APPTTM_MOD, vbInformation)

     End If


  End If

    '**************************************************************

End If
End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub


Private Sub lbl_aptdt_Click()
lbl_aptdt_Change
End Sub

Private Sub lbl_modalitycode_Change()

'If new_view_flg = 1 Then
'  If Len(Trim(lbl_modalitycode.Caption)) <> 0 Then
'  '***SERVICE NON AVAILABILITY********
'  f1
'
'  fillcal_sernonavail
'  f2
'  End If
'End If
End Sub

Private Sub msf_calendar_Click()
On Error GoTo MSG:
Dim var_month As String
Dim var_day As String
Dim i As Integer
zpublicHoliday = False
flg_sernonavail = False
Dim var_deptcode As String
Dim RSEQ As rdoResultset

For i = 0 To 12
If cbo_month.Text = cbo_month.List(i) Then
    If i < 9 Then  '** If i < 10 Then '*****5/7/2001 10changed to 9 as '010' appearing when october selected
        var_month = "0" & (i + 1)
    Else
        var_month = i + 1
    End If
End If
Next

If Len(Trim(msf_calendar.Text)) <> 0 Then

If msf_calendar.Text < 10 Then
    var_day = "0" & msf_calendar.Text
Else
    var_day = msf_calendar.Text
End If

End If
'****
Dim ST
ST = UCase(Format(var_day, "DDDD"))
    If (ST = "MONDAY" Or ST = "TUESDAY" Or ST = "FRIDAY" Or ST = "SATARDAY") And Trim(lbl_Modalitycode.Caption) = "ST" Then
        MsgBox " You Can Give Appointments only on Wednesday or Thursday "
    End If
    If Trim(lbl_Modalitycode.Caption) = "ST" Then
        timeslotno = 1
        strtime = 15
        cbo_apttm.clear
        While timeslotno < 9
            If strtime < 10 Then
            strtime = "0" & strtime
            End If
            cbo_apttm.AddItem strtime & ":00"
            cbo_apttm.AddItem strtime & ":20"
            cbo_apttm.AddItem strtime & ":40"
            strtime = Val(strtime) + 1
            timeslotno = timeslotno + 3
        Wend
    End If
    
'****
'If new_view_flg <> 2 Then
 If msf_calendar.CellBackColor <> &H8000000F And msf_calendar.CellBackColor <> &H80000010 Then
    lbl_aptdt.Caption = var_day & "/" & var_month & "/" & cbo_year.Text
 Else
    lbl_aptdt.Caption = ""
    lbl_aptdt.Caption = var_day & "/" & var_month & "/" & cbo_year.Text
    '25/01/2002-Parag
    'Boolean Value set to true if Selected date matches with Public Holiday Date
    If flg_sernonavail = True Then
        MsgBox "Service Will Not Be Available On This Day"
    Else
        MsgBox "On this day,our Hospital is CLOSED."
    End If
 End If
'End If



 '*************9/7/2001*******to show no of appointments in time slot box************
If Len(Trim(lbl_aptdt.Caption)) = 0 Then
    Exit Sub
End If

 refdate = GET_REVDATE(lbl_aptdt.Caption)
 refdateft = DateAdd("d", -3, refdate)
 refdatett = DateAdd("d", 3, refdate)
 refdatef = GET_REVDATE(CStr(refdateft))
 refdatet = GET_REVDATE(CStr(refdatett))

 str = "SELECT APNTDT,APNTTM,MODALITYCD,ser_det FROM RISLIB.APT_PATDET WHERE APNTDT between  '" & refdatef & "' and '" & refdatet & "' and CASENO = '" & tmh_ris_appointment.msk_caseno.FormattedText & "' and  dept_code = '" & "LZZ" & "' order by APNTDT, APNTTM"
 Set RSEQ = rdb.OpenResultset(str, 1, 1)
 If RSEQ.RowCount = 6 Then
 MsgBox "There are appointments Before Three Days and After Three Days, Please checkup HELP BUTTON"
 End If

 Dim rs_apptmcnt As rdoResultset

 Dim jind As Integer
 jind = InStr(cbo_deptname.Text, "$")
 var_deptcode = Trim(Mid(cbo_deptname.Text, jind + 1))

 On Error Resume Next
 str = "select  APNTTM, count(APNTTM)  from rislib.RREQUISITN where APNTDT = '" & GET_REVDATE(lbl_aptdt.Caption) & "' and MODALITYCD='" & tmh_ris_appointment.lbl_Modalitycode & "' and dept_code = '" & var_deptcode & "' group by APNTTM" '**29052001 added
 Set rs_apptmcnt = rdb.OpenResultset(str, 1, 3)

    'Dim i As Integer
    i = 0
  '*******COUNTING NO. OF APPTS IN THE SAME TIME SLOT  ********
  '******tms_arr CONTAINS TIME SLOT VS NO. OF APPTS FOR THAT SLOT*****

  If rs_apptmcnt.RowCount > 0 Then
    ReDim tms_arr(rs_apptmcnt.RowCount, 2) As String
    While Not rs_apptmcnt.EOF
        If IsNull(rs_apptmcnt(0)) = False Then
        tms_arr(i, 0) = Format(rs_apptmcnt(0), "hh:mm")
        tms_arr(i, 1) = rs_apptmcnt(1)
        Else
        tms_arr(i, 0) = ""
        tms_arr(i, 1) = rs_apptmcnt(1)
        End If

        rs_apptmcnt.MoveNext
        i = i + 1
    Wend
''    '*********filling appt time combo**********
''    Dim strtime
''    Dim timeslotno  As Integer
''
    cbo_apttm.clear
    If Trim(lbl_Modalitycode.Caption) = "MM" Or Trim(lbl_Modalitycode.Caption) = "MR" Or Trim(lbl_Modalitycode.Caption) = "CT" Or Trim(lbl_Modalitycode.Caption) = "CR" Then
        timeslotno = 1
        strtime = 8
        While timeslotno < 25
            If strtime < 10 Then
            strtime = "0" & strtime
            End If
            cbo_apttm.AddItem strtime & ":00"
            cbo_apttm.AddItem strtime & ":30"
            strtime = Val(strtime) + 1
            timeslotno = timeslotno + 2
        Wend
    End If
    If Trim(lbl_Modalitycode.Caption) = "2D" Or Trim(lbl_Modalitycode.Caption) = "PF" Or Trim(lbl_Modalitycode.Caption) = "ST" Then
        timeslotno = 1
        strtime = 9
        While timeslotno < 20
            If strtime < 10 Then
            strtime = "0" & strtime
            End If
            cbo_apttm.AddItem strtime & ":00"
            cbo_apttm.AddItem strtime & ":20"
            cbo_apttm.AddItem strtime & ":40"
            strtime = Val(strtime) + 1
            timeslotno = timeslotno + 3
        Wend
    End If
''
''    '********comparing each time slot in apttime combo with each element in tms_array********
    Dim no_ofappt As Integer
    Dim hour_var As String

    no_ofappt = rs_apptmcnt.RowCount

    For j = 0 To no_ofappt
        For i = 0 To cbo_apttm.ListCount - 1
            'jind = InStr(tms_arr(j, 0), ":")
            If Len(Trim(tms_arr(j, 1))) = 0 Then '***if no appts then 0 in timeslot array
               Exit Sub
            End If
            'hour_var = Mid(tms_arr(j, 0), 1, jind - 1)
            'If Left(hour_var, 1) = 0 Then '*****removing zero in front of time in time slot array for comparision purpose***
               'tms_arr(j, 0) = Mid(tms_arr(j, 0), 2)
            'End If
            If cbo_apttm.List(i) = tms_arr(j, 0) Then
               cbo_apttm.List(i) = cbo_apttm.List(i) & "             $ " & tms_arr(j, 1)
               GoTo lbl1:
            End If

        Next i
lbl1:     Next j

End If

 '****************************
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub msk_caseno_GotFocus()
    lbl_msg.Caption = "enter case number for which appointment is to be given"
End Sub

Private Sub msk_caseno_LostFocus()
'retrieve
On Error GoTo MSG:
If Len(Trim(msk_caseno.Text)) <> 0 Then
retrieve
    'If new_view_flg = 1 Then'**com 29052001
    If new_view_flg = 1 Or new_view_flg = 2 Then '***add 29052001
         If var_nopendingreq <> 1 Then
            frm_key.Enabled = True
            cbo_reqno.Enabled = True
           ' SendKeys "^{f4}"
            If new_view_flg = 1 Then
                lbl_msg.Caption = "select requisition number for which appointment is to be given"
            ElseIf new_view_flg = 2 Then
                lbl_msg.Caption = "select requisition number whose appointment is to be viewed"
            End If
            cbo_reqno.SetFocus
          Else
            clear_fields
            cmdbtn_vis ("I")
            frm_key.Enabled = False
            'atk
            'frm_entry.Enabled = False
         End If

'***29052001 comm below
'ElseIf new_view_flg = 2 Then
'        If msk_REQNNO.Enabled = True Then
'           lbl_msg.Caption = "enter requisition number for which appointment is to be given"
'            msk_REQNNO.SetFocus
'        ElseIf cmd_cancel.Enabled = True Then
'            cmd_cancel.SetFocus
'        ElseIf cmd_view.Enabled = True Then
'            cmd_view.SetFocus
'        End If
'****comm 29052001 till here

    End If


Else
clear_fields
cmdbtn_vis ("I")
frm_key.Enabled = False
'atk
'frm_entry.Enabled = False
'cbo_deptname.Enabled = False
'msk_caseno.Enabled = False
End If


Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Sub mskreplacedcbo()
On Error GoTo MSG:
Dim REQNNO_pad As String
'If Len(Trim(msk_REQNNO.Text)) <> 0 Then
If Len(Trim(cbo_reqno.Text)) <> 0 Then
    'REQNNO_pad = padzero(6, "Q", msk_REQNNO.Text, 8)
    'msk_REQNNO.Text = REQNNO_pad
    REQNNO_pad = req_chglen(cbo_reqno.Text)
    MSK_REQNUM = REQNNO_pad
    REQNUM = REQNNO_pad  'req_chglen(msk_REQNNO.Text)
    RREQUISITN_retrieve

    If var_aptnotgiven = 1 Then
        cmdbtn_vis ("I")
    End If
Else
    clear_fields
    If var_aptnotgiven = 1 Then
        cmdbtn_vis ("I")
        cmd_view.SetFocus
    Else
        If cmd_cancel.Enabled = True Then
            cmd_cancel.SetFocus
        End If
    End If

End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub


'Private Sub msk_REQNNO_LostFocus()
'Dim REQNNO_pad As String
'If Len(Trim(msk_reqnno.Text)) <> 0 Then
'    REQNNO_pad = padzero(6, "Q", msk_reqnno.Text, 8)
'    msk_reqnno.Text = REQNNO_pad
'    MSK_REQNUM = REQNNO_pad
'    REQNUM = REQNNO_pad  'req_chglen(msk_REQNNO.Text)
'    RREQUISITN_retrieve
'
'    If var_aptnotgiven = 1 Then
'        cmdbtn_vis ("I")
'    End If
'Else
'    clear_fields
'    cmdbtn_vis ("I")
'    cmd_view.SetFocus
'End If
'End Sub


Sub reqnextstg_chk()
On Error GoTo MSG:
reqnextstg_flg = 0

If new_view_flg = 1 Then
    rq_RREQUISITN(0) = req_chglen(cbo_reqno.Text)
Else
    rq_RREQUISITN(0) = req_chglen(cbo_reqno.Text) 'msk_REQNNO.Text '**29052001
End If
If flg_RREQUISITN = False Then
    Set rs_RREQUISITN = rq_RREQUISITN.OpenResultset(1, 1)
    flg_RREQUISITN = True
Else
    rs_RREQUISITN.MoreResults
    rs_RREQUISITN.Requery
End If

Dim K As Integer
If rs_RREQUISITN.RowCount > 0 Then
    'If Len(Trim(rs_RREQUISITN("REQNSTAT"))) = 0 Or IsNull(rs_RREQUISITN("REQNSTAT")) = True Then
    If Trim(rs_RREQUISITN("REQNSTAT")) = "A" Then
      reqnextstg_flg = 0
    ElseIf Trim(rs_RREQUISITN("REQNSTAT")) = "M" Then '****memo already generated'???//
        reqnextstg_flg = 2 '???// for memo
        K = MsgBox("        MEMO ALREADY GENERATED!" & Chr(13) & "You can delete the Appointment and Issue a fresh Appointment", vbInformation) '???//
    Else
     K = MsgBox("Requisition proceeded to next stage" & Chr(13) & "Deletion not possible!", vbCritical)
     reqnextstg_flg = 1
     cmdbtn_vis ("I")
     'atk
     'frm_entry.Enabled = False
    End If
End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub



Sub assign()

On Error GoTo MSG_err:
rdb.BeginTrans

'**********10/7/2001****
Dim k_ind As Integer
Dim var_apttm As String
k_ind = InStr(cbo_apttm.Text, "$")
If k_ind > 0 Then
var_apttm = Trim(Mid(cbo_apttm.Text, 1, k_ind - 1))
Else
var_apttm = Trim(cbo_apttm.Text)
End If
'******************till here 10/7/2001

 'If new_view_flg = 1 Then  '***new entry
 '********updating RREQUISITN table********************
    rq_RREQUISITNApptUpdt_Qr(0) = GET_REVDATE(lbl_aptdt.Caption)
 ' rq_RREQUISITNApptUpdt_Qr(1) = Format(cbo_apttm.Text, "hh:mm AM/PM") '****10/7/2001**comm
    rq_RREQUISITNApptUpdt_Qr(1) = Format(var_apttm, "hh:mm:ss AM/PM") '***10/7/2001 ****Add
    rq_RREQUISITNApptUpdt_Qr(2) = Trim(cbo_location.Text)   'Trim(txt_location.Text)'**changed from txt to cbo on 5/7/2001
    NO_CHARGE
    If No_Chargeflag = True Then
            rq_RREQUISITNApptUpdt_Qr(3) = "P"
    Else
        If MEMOGEN_VAR = 1 Then
            rq_RREQUISITNApptUpdt_Qr(3) = "M"
        Else
            rq_RREQUISITNApptUpdt_Qr(3) = "A"
        End If
    End If
    rq_RREQUISITNApptUpdt_Qr(4) = req_chglen(cbo_reqno.Text)
    rq_RREQUISITNApptUpdt_Qr.Execute 'UPDATE RREQUISITN
    '******************************
    If Trim(var_txtinstruct) <> Trim(txt_instructions.Text) Then
        rq_AppInstrInsrt_Qr(0) = req_chglen(cbo_reqno.Text)
        rq_AppInstrInsrt_Qr(1) = GET_REVDATE(lbl_aptdt.Caption)
        rq_AppInstrInsrt_Qr(2) = Null
        'MsgBox Len(txt_instructions.Text)
        rq_AppInstrInsrt_Qr(2) = Left(txt_instructions.Text, Len(txt_instructions.Text))
        rq_AppInstrInsrt_Qr.Execute
    End If

 ' inserting into patoolslib.ta102

       Insrt_Stmt(0) = msk_caseno.FormattedText
       Insrt_Stmt(1) = GET_REVDATE(lbl_aptdt.Caption)
       Insrt_Stmt(2) = Format(var_apttm, "hh:mm:ss AM/PM")
       Insrt_Stmt(3) = Trim(lbl_proccode.Caption)
       Insrt_Stmt(4) = "By RIS"
       Insrt_Stmt(5) = "-"
       Insrt_Stmt(6) = "-"
       Insrt_Stmt(7) = "-"
       Insrt_Stmt(8) = "-"
       rq_RREQUISITN(0) = REQNUM
       rs_RREQUISITN.MoreResults
       rs_RREQUISITN.Requery
       'MsgBox rs_RREQUISITN.RowCount
       Insrt_Stmt(9) = rs_RREQUISITN("refbydoc")

       Insrt_Stmt(10) = Trim(lbl_refby.Caption)
       Insrt_Stmt(11) = GET_REVDATE("01/01/2000")
       Insrt_Stmt(12) = GET_REVDATE(msk_sysdate.Text)
       Insrt_Stmt.Execute
    rdb.CommitTrans
    i = MsgBox("Record successfully Added / Updated", vbInformation + vbOKOnly)
    cmdbtn_vis ("S")
    frm_key.Enabled = False
Exit Sub

MSG_err:
    rdb.RollbackTrans
    MsgBox Error(Err)

End Sub


'_________________




Private Sub cbo_month_Click()

m = cbo_month.ListIndex + 1
f1
f2
End Sub




Private Sub cbo_year_Click()

Y = cbo_year

f1
f2
End Sub

Public Sub f1()
On Error GoTo MSG:
D1 = Y & "/" & m & "/1"
If m = 12 Then
D2 = (Y + 1) & "/" & "1/1"
Else
D2 = Y & "/" & (m + 1) & "/1"
End If

Dim firstday As Integer, totdays As Integer
firstday = Weekday(D1)
totdays = DateDiff("d", D1, D2)

'clear the calendar
For i = 1 To 6
For j = 0 To 6
msf_calendar.row = i
msf_calendar.col = j
msf_calendar.TextMatrix(i, j) = ""
msf_calendar.CellBackColor = &H8000000E
Next
Next

'mark sundays with red color
'For i = 1 To 6
'msf_calendar.Row = i
'msf_calendar.Col = 0
'msf_calendar.CellBackColor = &H8000000F '&HFF&
'Next

'mark sundays and saturdays with red color
For i = 1 To 6
For j = 0 To 6
msf_calendar.row = i
msf_calendar.col = j
If msf_calendar.col = 0 Then
msf_calendar.CellBackColor = &H8000000F '&H8000000F&'&HFF&
End If
Next
Next

'***fill the date numbers in the grid******
msf_calendar.row = 1
msf_calendar.col = firstday - 1
For i = 1 To totdays
    msf_calendar.Text = i
        If msf_calendar.col = 6 Then
            msf_calendar.col = 0
            msf_calendar.row = msf_calendar.row + 1
        Else
            msf_calendar.col = msf_calendar.col + 1
        End If

Next
'****mark the holidays***********
Dim r As Integer, c As Integer
For i = 0 To totholidays - 1
        If Year(hlist(i)) = Y And month(hlist(i)) = m Then
            j = Day(hlist(i))
            For r = 1 To 6
            For c = 0 To 6
            msf_calendar.row = r
            msf_calendar.col = c
            If Val(msf_calendar.Text) = j Then
                msf_calendar.CellBackColor = &H8000000F  '&HFF&
            End If
            Next
            Next
            End If
Next

'lbl_monthyear.Caption = cbo_month.Text & cbo_year.Text'**comm 4/7/2001
lbl_monthyear.Caption = "" '**  add 4/7/2001
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Sub f2()
On Error GoTo MSG:
D1 = Y & "/" & m & "/1"  '***THIS MONTH
If m = 12 Then
D2 = (Y + 1) & "/" & "1/1" '*******NEXT MONTH START DATE
Else
D2 = Y & "/" & (m + 1) & "/1" '*******NEXT MONTH START DATE
End If

Dim firstday As Integer, totdays As Integer
firstday = Weekday(D1)
totdays = DateDiff("d", D1, D2)

Dim S1 As String, S2 As String
Dim FROMYEARMONTH As String, TOYEARMONTH As String, CURRYEARMONTH As String
For i = 0 To NVLROWS - 1
    FROMYEARMONTH = Year(GET_REVDATE(fromdt(i))) & Format(GET_REVDATE(fromdt(i)), "mm")
    TOYEARMONTH = Year(GET_REVDATE(todt(i))) & Format(GET_REVDATE(todt(i)), "mm")
    CURRYEARMONTH = Y & Format(D1, "mm")

If Val(CURRYEARMONTH) >= Val(FROMYEARMONTH) And Val(CURRYEARMONTH) <= Val(TOYEARMONTH) Then
    If GET_REVDATE(Format(D1, "dd/mm/yyyy")) > GET_REVDATE(fromdt(i)) Then
        S1 = D1
    Else
        S1 = fromdt(i)
    End If

    If GET_REVDATE(Format(D2, "dd/mm/yyyy")) >= GET_REVDATE(todt(i)) Then
    S2 = todt(i)
    Else
    S2 = D2
    End If

    Dim nonavldays As Integer, strdt As Integer
    Dim DTNEW1
    Dim DTNEW2
    DTNEW1 = GET_REVDATE(S1)
    DTNEW2 = GET_REVDATE(S2)

    'nonavldays = DateDiff("d", S1, s2)
    nonavldays = DateDiff("D", DTNEW1, DTNEW2)
    'strdt = Format(GET_REVDATE(Format(S1, "dd/mm/yyyy")), "dd")
    strdt = Format(GET_REVDATE(S1), "dd")
    msf_calendar.row = 1
    msf_calendar.col = 0
    While Val(msf_calendar.Text) <> Val(strdt)
        If msf_calendar.col = 6 Then
            msf_calendar.row = msf_calendar.row + 1
            msf_calendar.col = 0
        Else
            msf_calendar.col = msf_calendar.col + 1
        End If
    Wend

    For j = 0 To nonavldays
        msf_calendar.CellBackColor = &H80000010 '&H8000000F
        If msf_calendar.col = 6 Then
            msf_calendar.row = msf_calendar.row + 1
            msf_calendar.col = 0
        Else
            msf_calendar.col = msf_calendar.col + 1
        End If
    Next

End If

Next




'Dim r As Integer, c As Integer
'For i = 0 To totnonavldays '- 1
'        If Year(srvnonavail_list(i)) = Y And Month(srvnonavail_list(i)) = m Then
'            j = Day(srvnonavail_list(i))
'            For r = 1 To 6
'            For c = 0 To 6
'            msf_calendar.Row = r
'            msf_calendar.Col = c
'            If Val(msf_calendar.Text) = j Then
'                msf_calendar.CellBackColor = &H8000000F  '&HFF&
'            End If
'            Next
'            Next
'            End If
'Next
'
'lbl_monthyear.Caption = cbo_month.Text & cbo_year.Text
Exit Sub
MSG:
MsgBox Error(Err)
End Sub



Private Sub txt_instructions_GotFocus()
txt_instructions.Top = txt_instructions.Top - 500
txt_instructions.Width = txt_instructions.Width + 500
txt_instructions.Height = txt_instructions.Height + 500

End Sub

Private Sub txt_instructions_KeyPress(KeyAscii As Integer)
If KeyAscii = 36 Or KeyAscii = 39 Or KeyAscii = 124 Then
KeyAscii = 0
End If

End Sub

'Private Sub txt_location_KeyPress(KeyAscii As Integer) '****comm 5/7/2001
'If KeyAscii = 36 Or KeyAscii = 39 Or KeyAscii = 124 Then
'KeyAscii = 0
'End If
'
'End Sub
'*************
Private Sub txt_instructions_LostFocus()
txt_instructions.Top = txt_instructions.Top + 500
txt_instructions.Width = txt_instructions.Width - 500
txt_instructions.Height = txt_instructions.Height - 500


End Sub

'Public Sub cmdnew()
'
'    If RISAUTH_RS("NEW") = 1 Then
'        CMD_NEW.Enabled = True
'    Else
'        CMD_NEW.Enabled = False
'
'    End If
'End Sub
'Public Sub cmdview()
'
'    If RISAUTH_RS("VIEW") = 1 Then
'        cmd_view.Enabled = True
'
'    Else
'        cmd_view.Enabled = False
'    End If
'End Sub
'Public Sub CMDDELETE()
'
'    If RISAUTH_RS("DELETE") = 1 Then
'    cmd_delete.Enabled = True
'    Else
'    cmd_delete.Enabled = False
'    End If
'
'End Sub
'Public Sub cmdsave()
'
'    If RISAUTH_RS("SAVE") = 1 Then
'    cmd_save.Enabled = True
'    Else
'    cmd_save.Enabled = False
'    End If
'
'End Sub
'
'Public Sub CMDPRINT()
'
'    If RISAUTH_RS("PRINT") = 1 Then
'    cmd_print.Enabled = True
'
'    Else
'    cmd_print.Enabled = False
'    End If
'
'End Sub
'Public Sub CMDQUERY()
'
'    If RISAUTH_RS("QUERY") = 1 Then
'    cmd_query.Enabled = True
'    Else
'    cmd_query.Enabled = False
'    End If
'
'End Sub
'Public Sub CMDHELP()
'
'    If RISAUTH_RS("HELP") = 1 Then
'    cmd_help.Enabled = True
'    Else
'    cmd_help.Enabled = False
'    End If
'
'End Sub
'Public Sub CMDDIARY()
'
'    If RISAUTH_RS("DIARY") = 1 Then
'    CMD_DIALY.Enabled = True
'    Else
'    CMD_DIALY.Enabled = False
'    End If
'
'End Sub
'
