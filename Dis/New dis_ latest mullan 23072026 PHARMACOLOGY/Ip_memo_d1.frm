VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form ip_memo_d1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   9120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox TXTNET 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      TabIndex        =   64
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TXTPRTTOT 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   62
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport ipcr 
      Left            =   9120
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   9120
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   1692
      Left            =   120
      Negotiate       =   -1  'True
      OleObjectBlob   =   "Ip_memo_d1.frx":0000
      TabIndex        =   58
      Top             =   3720
      Width           =   9372
   End
   Begin VB.CommandButton Back 
      Caption         =   "&Back"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6720
      Width           =   975
   End
   Begin VB.Frame pat_det 
      Caption         =   "PATIENT _ DETAILS"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   9375
      Begin VB.Label lbl_deposit 
         Alignment       =   2  'Center
         BackColor       =   &H00C000C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   67
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lbl_smartcard 
         BackColor       =   &H00C000C0&
         Caption         =   "Smart Card"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   66
         Top             =   960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl_tr_no 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   3240
         TabIndex        =   39
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label lbl_cr_card_no 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3240
         TabIndex        =   38
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Credit_card No."
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lbl_cat 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   36
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lbl_age 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8400
         TabIndex        =   35
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lbl_sex 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   7440
         TabIndex        =   34
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lbl_name 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   1920
         TabIndex        =   33
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label lbl_age1 
         Caption         =   "Age"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8040
         TabIndex        =   32
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lbl_cat1 
         Caption         =   "Category"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbl_sex1 
         Caption         =   "Sex"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7080
         TabIndex        =   30
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lbl_name1 
         Caption         =   "Patient's Name"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   600
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl_yrs 
         Caption         =   "Yrs."
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   8880
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Ward/ Floor :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4440
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lbl_wf_bed 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5640
         TabIndex        =   26
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label17 
         Caption         =   "Trust A/C No."
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Type :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4680
         TabIndex        =   24
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbl_typ 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5640
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.CommandButton New1 
      Cancel          =   -1  'True
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1560
      TabIndex        =   0
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox txt_doc_sec_name 
      Height          =   225
      Left            =   7320
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton view 
      Caption         =   "&View"
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
      Left            =   5160
      TabIndex        =   21
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton print1 
      Caption         =   "&Print"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton exit1 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7560
      TabIndex        =   17
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton delete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      TabIndex        =   20
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton commit 
      Caption         =   "&Save"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox txt_rem 
      Height          =   288
      Left            =   1080
      TabIndex        =   13
      Top             =   6240
      Width           =   4935
   End
   Begin VB.TextBox Text22 
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
      Left            =   1200
      TabIndex        =   19
      Top             =   7200
      Width           =   8055
   End
   Begin VB.ComboBox Com_dno 
      BackColor       =   &H8000000E&
      ForeColor       =   &H80000012&
      Height          =   288
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   3732
   End
   Begin VB.CommandButton cancel 
      Caption         =   "&Cancel"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6720
      Width           =   1095
   End
   Begin MSMask.MaskEdBox me_cno 
      Height          =   372
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   1452
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      HideSelection   =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">??/#####"
      PromptChar      =   "_"
   End
   Begin MSFlexGridLib.MSFlexGrid flex1 
      Height          =   1692
      Left            =   480
      TabIndex        =   10
      Top             =   3720
      Width           =   7812
      _ExtentX        =   13785
      _ExtentY        =   2990
      _Version        =   393216
      FixedCols       =   0
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      GridLines       =   3
      AllowUserResizing=   3
      MousePointer    =   4
      FormatString    =   "^          CODE            |                          SERVICES/PROCEDURES/INVESTIGATIONS"
   End
   Begin MSMask.MaskEdBox me_mno 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      HideSelection   =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">??######"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox me_date 
      Height          =   375
      Left            =   7680
      TabIndex        =   47
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   128
      ForeColor       =   -2147483634
      AutoTab         =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox me_rdate 
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483627
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      HideSelection   =   0   'False
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
   Begin MSFlexGridLib.MSFlexGrid flex2 
      Height          =   1695
      Left            =   2880
      TabIndex        =   11
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2990
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      WordWrap        =   -1  'True
      GridLines       =   3
      SelectionMode   =   1
      AllowUserResizing=   3
      MousePointer    =   4
      FormatString    =   "^     DOCTOR NAME"
   End
   Begin MSMask.MaskEdBox M_number 
      Height          =   372
      Left            =   1320
      TabIndex        =   9
      Top             =   600
      Width           =   1812
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   8404992
      ForeColor       =   -2147483643
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">??######"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txt_unit_doc_name 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5760
      TabIndex        =   5
      Text            =   " "
      Top             =   2880
      Width           =   2655
   End
   Begin VB.ComboBox d_u_name 
      Height          =   315
      Left            =   5760
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2895
   End
   Begin VB.ComboBox cbo_reqno 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   3240
      Width           =   3495
   End
   Begin VB.TextBox txt_req_no 
      Height          =   315
      Left            =   1920
      TabIndex        =   59
      Top             =   3240
      Width           =   3315
   End
   Begin VB.Label Label7 
      Caption         =   "protocol patient"
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6000
      TabIndex        =   65
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label19 
      Caption         =   "NET PAYABLE AMOUNT :"
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4560
      TabIndex        =   63
      Top             =   5880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label20 
      Caption         =   "LESS PROTOCOL AMOUNT:"
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   61
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "&Entered By :"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6120
      TabIndex        =   60
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label23 
      Caption         =   "REQUISITION DATE :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5640
      TabIndex        =   55
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   9480
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label lbl_cno 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1320
      TabIndex        =   54
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DEPT. CODE  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   6000
      TabIndex        =   53
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "DOCTOR/UNIT NAME :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3240
      TabIndex        =   52
      Top             =   2880
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label11 
      Caption         =   "DEPARTMENT :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   960
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      Caption         =   "CASE NO :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   50
      Top             =   1200
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "TOTAL :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4560
      TabIndex        =   49
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lbl_tot_amt 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000080&
      Height          =   372
      Left            =   5640
      TabIndex        =   12
      Top             =   5400
      Width           =   1572
   End
   Begin VB.Label lbl_dcode 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lbl_dname 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label lbl_doc_code 
      Caption         =   "DOCTOR'S CODE  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   2880
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      Height          =   7560
      Left            =   60
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "IN-PATIENT MEMO OF CHARGES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   46
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000B&
      Caption         =   "Message :"
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
      Left            =   120
      TabIndex        =   45
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
      Caption         =   "MEMO NO :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   720
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DATE :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6000
      TabIndex        =   43
      Top             =   720
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9480
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "REQUISITOIN No. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   3240
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_dno 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7680
      TabIndex        =   41
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Left            =   120
      TabIndex        =   56
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "ip_memo_d1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'31/03/2003-Parag/Sandhya
'Mapping Patch w.f.to New Schedule of Charges APR-2003.
Dim ZUSESERCODEMAP As Boolean
Dim ZUSESERINDMAP As Boolean

Dim zsercodemap As rdoQuery
Dim zrssercodemap As rdoResultset

Dim zserindmap As rdoQuery
Dim zrsserindmap As rdoResultset

Dim zstr As String

Dim RQ_SERIES As rdoQuery
Dim RS_SERIES As rdoResultset
Dim RS_SERIESflag As Boolean
Dim K As String                     ' ADDED 11/11/2001
Dim k1 As String                    ' ADDED 11/11/2001
Dim S
Dim REQGOTFLAG As Boolean
Dim rq_updtSERIES As rdoQuery
Dim rq_reqdet As rdoQuery
Dim RS_reqdet As rdoResultset
Dim RS_reqdetflag As Boolean
Dim rq_updtreqdet As rdoQuery
Dim rq_updtdelreqdet As rdoQuery
Dim RQ_REQDETREQNO As rdoQuery
Dim RS_REQDETREQNO As rdoResultset
Dim RS_REQDETREQNOflag As Boolean
Dim requisitn_qr As rdoQuery
Dim requisitn_rs As rdoResultset
Dim requisitn_rsflag As Boolean
Dim UPDTA107_QR As rdoQuery
Dim InPt_ChkQr As rdoQuery
Dim RSINPT As rdoResultset
Dim RSINPTflag As Boolean
Dim DATE11 As Date
Dim REQUISITN_UPDT As rdoQuery
Dim RQ_reqdet1 As rdoQuery
Dim rs_reqdet1 As rdoResultset
Dim RS_reqdet1flag As Boolean
''''FLEX1.CLICK START FROM HERE
Dim REQNO As String
'''PARAS

Dim Exist_ChkQr As rdoQuery
Dim rs22_CHK As rdoResultset
Dim rs22_CHKflag As Boolean

'''PARAS

Dim invcat As Boolean
Dim NOCUR As Integer
Dim EmptyGrid '***28499r
Dim CASENO As String
Dim AmtNotEq As Integer '**19599
Dim flag1 As Integer
Dim sql As String  'BILLOPT 61299 ADDED
Dim RS_LVLQR As rdoQuery 'BILLOPT 61299 ADDED
Dim RS12QR As rdoQuery  'BILLOPT 61299 ADDED
Dim RS14QR As rdoQuery 'BILLOPT  61299 ADDED
Dim RS16QR As rdoQuery 'BILLOPT 61299 ADDED
Dim RS15QR As rdoQuery 'BILLOPT 61299 ADDED
Dim RS2QR As rdoQuery 'BILLOPT 61299 ADDED
Dim RS33QR As rdoQuery 'BILLOPT 61299 ADDED

Dim RS33QRCD As rdoQuery

Dim RS21QR As rdoQuery 'BILLOPT 61299 ADDED
Dim RS18QR As rdoQuery 'BILLOPT 61299 ADDED
Dim RSDBTPQR As rdoQuery 'BILLOPT 61299 ADDED
Dim RSDNAMEQR As rdoQuery 'BILLOPT 61299 ADDED
Dim RS5QR As rdoQuery 'BILLOPT 61299 ADDED
Dim Ma104_Qr As rdoQuery
Dim ma104_rs As rdoResultset
Dim ma104_rsflag As Boolean
Dim RS_CR_QR As rdoQuery 'BILLOPT 61299 ADDED
Dim RS_TR_QR As rdoQuery 'BILLOPT 61299 ADDED
Dim RS44QR As rdoQuery
Dim addta107_QR As rdoQuery 'BILLOPT 61299 ADDED
Dim addta1071_QR As rdoQuery 'BILLOPT 61299 ADDED

Dim rs7qr As rdoQuery 'BILLOPT 61299 ADDED
Dim RS10QR As rdoQuery 'BILLOPT 61299 ADDED
Dim rs22qr As rdoQuery 'BILLOPT ADD

 Dim RS15QR1 As rdoQuery
 Dim RS151  As rdoResultset
 Dim RS151flag  As Boolean



Dim rs_cr As rdoResultset
Dim rs_crflag As Boolean
Dim rs_tr As rdoResultset
Dim rs_trflag As Boolean
Dim rs30 As rdoResultset
Dim rs30flag As Boolean
Dim rs31 As rdoResultset
Dim rs31flag As Boolean
Dim rs1 As rdoResultset
Dim rs1flag As Boolean
Dim rs33 As rdoResultset
Dim rs33flag As Boolean
Dim rs22 As rdoResultset
Dim rs22flag As Boolean
Dim rs3 As rdoResultset
Dim rs3flag As Boolean
Dim rs4 As rdoResultset
Dim rs4flag As Boolean
Dim rs5 As rdoResultset
Dim rs5flag As Boolean
Dim rs6 As rdoResultset
Dim rs6flag As Boolean
Dim rs7 As rdoResultset
Dim rs7flag As Boolean
Dim rs8 As rdoResultset
Dim rs8flag As Boolean
Dim rs10 As rdoResultset
Dim rs10flag As Boolean
Dim rs12 As rdoResultset
Dim rs12flag As Boolean
Dim RS11 As rdoResultset
Dim rs11flag As Boolean
Dim rs13 As rdoResultset
Dim rs13flag As Boolean
Dim RS15 As rdoResultset
Dim rs15flag As Boolean
Dim RS14 As rdoResultset
Dim rs14flag As Boolean
Dim RS16 As rdoResultset
Dim rs16flag As Boolean
Dim RS17 As rdoResultset
Dim rs17flag As Boolean
Dim rs2 As rdoResultset
Dim rs2flag As Boolean
Dim RS20 As rdoResultset
Dim rs20flag As Boolean
Dim RS18 As rdoResultset
Dim rs18flag As Boolean
Dim RS19 As rdoResultset
Dim rs19flag As Boolean
Dim RSDBTP As rdoResultset '24599
Dim rsdbtpflag As Boolean
Dim ncchg As String  ' added for z-series memos***4899
Dim rs_dname As rdoResultset
Dim rs_dnameflag As Boolean
Dim a As Integer
Dim RS As rdoResultset
Dim rsflag As Boolean
Dim rs21 As rdoResultset
Dim rs21flag As Boolean
Dim flag As Integer
Private mTotalRows& ' Contains the total rows in the set of records
Private UserData() As Variant ' 2-dimensional array containing records
Private Const MAXCOLS = 6 ' Maximum number of fields in record set.
Dim FLGTA As Integer '*17/12/98
Dim FLGCR As Integer '**17/12/98
Dim PATYPE As String '**17/12/98
Dim DOCFLAG As Boolean
Dim PrintResult%

Dim SELDMULPROCEQR As rdoQuery
Dim SELDMULPROCERS As rdoResultset
Dim SELDMULPROCEFLAG As Boolean

Dim SELDMOLEPROCQR As rdoQuery
Dim SELDMOLEPROCRS As rdoResultset
Dim SELDMOLEPROCFLAG As Boolean

Dim zgetservice As rdoQuery
Dim zrsgetservice As rdoResultset
Dim ZRSGETSERVICEFLAG1 As Boolean

Dim zmaptable As rdoQuery
Dim zrsmaptable As rdoResultset
Dim ZRSMAPTABLEFLAG As Boolean

Dim zupdtmap As rdoQuery

Dim zta1071 As rdoQuery
Dim zrsta1071 As rdoResultset
Dim Zrsta1071flag As Boolean

Dim selreq_qr As rdoQuery
Dim selreq_rs As rdoResultset

''*****FOR BP CROSSMATCH******
Dim SELH002_QR As rdoQuery
Dim SELH002_RS As rdoResultset
Dim SELH002_FLAG As Boolean
Dim dept As String
Dim SELCOLLECTION_QR As rdoQuery
Dim SELCOLLECTION_RS As rdoResultset
Dim SELCOLLECTION_FLAG As Boolean
''*****FOR BP CROSSMATCH******
Dim SEL_M3ICU_QR As rdoQuery
Dim SEL_M3ICU_RS As rdoResultset
Dim SEL_M3ICU_FLG As Boolean

'' BIJESH

Dim ZAS400DTTM As rdoResultset
Dim zgetSerCodes As rdoQuery
Dim zrsgetSercodes As rdoResultset
Dim memoduplicate As Boolean
Dim RSTA1071 As rdoResultset
'' END
''' 24.9.2008 - SANDHYA protocol patient starts
Dim PROTFLG As Boolean
Dim NPROTFLG As Boolean
Dim PROTCASE_QR As rdoQuery
Dim PROTCASE_RS As rdoResultset
Dim RSPROTCASE As rdoResultset
Dim PROTSERV_QR As rdoQuery
Dim PROTSERV_RS As rdoResultset
Dim PROTMEMO_Qr As rdoQuery
Dim PROTCASE_UPD As rdoQuery
Dim QR_PRTMEMO As rdoQuery
Dim RS_PRTMEMO As rdoResultset
Dim zupdprotmap As rdoQuery
Dim zPRTMEMOflag As rdoQuery
Dim PTOT As Long
Dim NTOT As Long
''protocol patient ends

'''' added by shraddha
'''' dated 16.8.2010 to check if service code exist in table ser_ref for duplication
Dim qrser_ref As rdoQuery
Dim rsser_ref As rdoResultset
'''' ADDED BY SANDHYA DATED 14.7.2011
'''' SMART CARD
Dim scard As Boolean
Dim PCARD As PISWriteCard.ClsWriteCard
Dim PINFO As PatientInformation
Dim AMT_RCVD As String
Dim SMCARDFLG As Boolean
Dim scard_bal_amt As Double
Dim FLGSM As Integer
''''' ADDED BY SANDHYA DATED 18.07.2012 FOR BP PATIENT RAJIV GANDHI
Dim INSBPMEMO As rdoQuery
'Dim BPPATFLAG As Boolean
Dim QRBPPATN As rdoQuery
Dim RS_BPPATN As rdoResultset
Dim UPD_bppatn As rdoQuery
Dim INSBPPATMEMO As rdoQuery

Sub ipmemocr()
Dim STR5 As String
Dim STR6 As String
Dim STR7 As String
Dim STR8 As String
Dim STR9 As String
Dim STR10 As String
Dim STR11 As String
Dim str12 As String
Dim CHKTOT As Double
Dim i As Integer
Dim j As Integer
Dim K As Integer
Dim samt As String
Dim l As Integer
Dim a As String
STR11 = COMP_IP
a = Chr(34)
samt = 0
i = 0
K = 0
ipcr.ReportFileName = "C:\PB121298\IPMEMO.rpt"
For j = 0 To 120 '261099
    ipcr.Formulas(j) = " "
Next j

If flag = 1 Then
ipcr.Formulas(0) = "fmemono = '" & me_mno.Text & "'"    '"M-" & me_mno.Text & "'"
ipcr.Formulas(3) = "fdeptname = '" & Com_dno.Text & "'"
ipcr.Formulas(2) = "fcaseno = '" & CASENO & "'"
str12 = CASENO
    If rs22_CHKflag = False Then
        Exist_ChkQr(0) = me_mno.Text   '"M-" & me_mno.Text
        Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
        rs22_CHKflag = True
    Else
        Exist_ChkQr(0) = me_mno.Text
        rs22_CHK.MoreResults
        rs22_CHK.Requery
    End If
STR10 = "*A*"
Else
ipcr.Formulas(0) = "fmemono = '" & M_number & "'"     '"M-" & M_number & "'"
ipcr.Formulas(2) = "fcaseno = '" & lbl_cno & "'"
ipcr.Formulas(3) = "fdeptname = '" & lbl_dname & "'"
     '"M-" & M_number
    str12 = lbl_cno
    If rs22_CHKflag = False Then
        Exist_ChkQr(0) = M_number
        Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
        rs22_CHKflag = True
    Else
        Exist_ChkQr(0) = M_number
        rs22_CHK.MoreResults
        rs22_CHK.Requery
    End If
STR10 = "*V*"
End If
 'If DOCFLAG = True Then
    ipcr.Formulas(1) = "doc_name = " & a & Trim(txt_unit_doc_name.Text) & a & ""
    ipcr.Formulas(8) = "doc_code = " & a & Trim(lbl_dcode.Caption) & a & ""
' ElseIf DOCFLAG = False Then
'    ipcr.Formulas(1) = "doc_name = " & A & A & ""
'    ipcr.Formulas(8) = "doc_code = " & A & 0 & A & ""
' End If

'IPCR.Formulas(2) = "Fcaseno = '" & CASENO & "'"

ipcr.Formulas(4) = "fdcode = " & a & lbl_dno & a & ""
 
ipcr.Formulas(5) = "NAME_CON = " & a & Trim(UCase(LBL_NAME)) & a & ""


ipcr.Formulas(6) = "fpat_catg = '" & lbl_cat & "'" 'modify
ipcr.Formulas(7) = "fward = '" & lbl_wf_bed & "'" 'modify
'IPCR.Formulas(8) = "acc_num = " & A & "lbl_tr_no" & A & ""
'IPCR.Formulas(9) = "acc_head = " & A & Trim(lbl_acc_head) & A & ""
ipcr.Formulas(9) = "freqno = " & a & Trim(txt_req_no.Text) & a & ""
ipcr.Formulas(10) = "ftot = " & a & " *****" & Format(lbl_tot_amt, "#,##,##,##,##0.00") & a & ""
ipcr.Formulas(11) = "freqdate = " & a & Trim(me_rdate.Text) & a & ""
ipcr.Formulas(12) = "fremark = " & a & Trim(txt_rem.Text) & a & ""
ipcr.Formulas(13) = "fmemodate  = " & a & Trim(me_date.Text) & a & ""
'STR1 = RTrim(username) 'change this var
 STR5 = RTrim(usr_name)
 STR6 = Format(DATE11, "DD/MM/YYYY")
 STR7 = time1
 'Exist_ChkQr(0) = "M-100010"
'Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
If rs22_CHK.RowCount <> 0 Then
 STR9 = Trim(rs22_CHK("D_S_NAME"))
End If
 'STR8 = Trim(STR5) + "  " + Trim(STR6) + "  " + Trim(STR7)
 STR8 = Trim(STR9) + " / " + Trim(STR5) + " " + Trim(STR10) + " " + " " + STR11 + " " + Trim(STR6) + "  " + Trim(STR7)
ipcr.Formulas(14) = "us_dt_tm = " & a & STR8 & a & ""

'''' ADDED BY SANDHYA DATED 18-7-2012 FOR RGJAY PATIENT
If BPPATFLAG = True Then
  If flag = 1 Then
    Set RS_BPPATN = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.BPPATIENT WHERE CASENO ='" & CASENO & "'", 1, 1)
  End If
  If flag = 2 Then
    Set RS_BPPATN = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.BPPATIENT WHERE CASENO ='" & lbl_cno & "'", 1, 1)
  End If
End If


 If rs2("PAT_TYPE") = "TA" Then
     ipcr.Formulas(15) = "cr_tr = '" & rs_tr("TR_ACC_NO") & "'"
     '''' ADDED BY SANDHYA DATED 18-7-2012 FOR RGJAY PATIENT
 ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = False Then
    ipcr.Formulas(15) = "cr_tr = '" & Trim(rs_cr("CR_CARD")) & "'"
    ipcr.Formulas(150) = "PATTP = '" & "COMPANY PATIENT" & "'"
 ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = True Then
    ipcr.Formulas(15) = "cr_tr = '" & Trim(RS_BPPATN("BPCARD_NO")) & "'"
    ipcr.Formulas(150) = "PATTP = '" & "RGJAY PATIENT" & "'"
   
  '''' ADDED BY SANDHYA FOR SMARTCARD PATIENT
   ElseIf rs2("PAT_TYPE") = "CP" And scard = True Then
    ipcr.Formulas(15) = "cr_tr = '" & Right(me_cno.Text, 5) & Left(me_cno.Text, 2) & "'"
    ipcr.Formulas(150) = "PATTP = '" & "SMART CARD PATIENT" & "'"
   ElseIf rs2("PAT_TYPE") = "CP" And scard = False Then
    ipcr.Formulas(15) = "cr_tr = '" & "    " & "'"
    ipcr.Formulas(150) = "PATTP = '" & "    " & "'"
 End If
    






i = 0
K = 0
 
If flag = 1 Then
    If rs7flag = False Then
        rs7qr(0) = me_mno.Text
        Set rs7 = rs7qr.OpenResultset(1, 3)
        rs7flag = True
    Else
        rs7qr(0) = me_mno.Text
        rs7.MoreResults
        rs7.Requery
    End If
Else
    If rs7flag = False Then
        rs7qr(0) = M_number
        Set rs7 = rs7qr.OpenResultset(1, 1)
        rs7flag = True
    Else
        rs7qr(0) = M_number
        rs7.MoreResults
        rs7.Requery
    End If
End If
     

Dim rrow

DBGrid1.Refresh

rrow1 = 0
If rs7.RowCount Then
        rs7.MoveFirst
        For rrow1 = 0 To rs7.RowCount - 1
  ipcr.Formulas(17 + rrow1 + K) = "s" & rrow1 & " = " & a & Trim(rs7(1).Value) & a & ""
                If Not IsNull(rs7(1).Value) Then
            If rs18flag = False Then
                RS18QR(0) = rs7(1).Value
                Set RS18 = RS18QR.OpenResultset(1, 3)
                rs18flag = True
            Else
                RS18QR(0) = rs7(1).Value
                RS18.MoreResults
                RS18.Requery
            End If
  ipcr.Formulas(40 + rrow1 + K) = "ser_det" & rrow1 & " = " & a & Trim(RS18("SER_DET").Value) & a & ""   'modify
  ipcr.Formulas(64 + rrow1 + K) = "q" & rrow1 & " = " & a & Trim(rs7("qty")) & a & ""   'modify
                samt = Format((rs7(2).Value), "#,##,##,##,###0.00")
                samt = RTrim(samt)
                CHKTOT = CHKTOT + CDbl(rs7(2))
                l = Len(samt)
                samt = Space(18 - l) + samt
        ipcr.Formulas(88 + rrow1 + K) = "ser_amt" & rrow1 & " = " & a & samt & a & "" 'modify
        ipcr.Formulas(124 + rrow1 + K) = "REQ" & rrow1 & " = " & a & rs7("REQ_NO").Value & a & ""  'modify
        End If
        K = K + 1
        rs7.MoveNext
       Next rrow1
    End If
If rs2flag = False Then
    RS2QR(0) = str12
    Set rs2 = RS2QR.OpenResultset(1, 3)
    rs2flag = True
Else
    RS2QR(0) = str12
    rs2.MoreResults
    rs2.Requery
End If
If rs2.RowCount > 0 Then
'MsgBox rs2.RowCount
    If requisitn_rsflag = False Then
        requisitn_qr(0) = str12
        Set requisitn_rs = requisitn_qr.OpenResultset(2, 3)
        requisitn_rsflag = True
    Else
        requisitn_qr(0) = str12
        requisitn_rs.MoreResults
        requisitn_rs.Requery
    End If
    'MsgBox requisitn_rs.RowCount
        If requisitn_rs.RowCount > 0 Then
            If Trim(rs2("pat_catg")) = Trim(requisitn_rs("pat_catg")) Then
            'IPCR.Formulas(15) = "cr_tr = '" & rs_tr("TR_ACC_NO") & "'"
            ipcr.Formulas(148) = "catchg = '" & "" & "'"  'modify
            Else
            ipcr.Formulas(148) = "catchg = '" & "Patient Category is changed :" & requisitn_rs("pat_catg") & "  to  " & rs2("pat_catg") & "'"  'modify
            End If
        End If
End If


If CDbl(lbl_tot_amt) = CDbl(CHKTOT) Then
'Parag
ipcr.CopiesToPrinter = 1
ipcr.WindowState = crptMaximized
ipcr.Destination = crptToPrinter
 'ipcr.Destination = crptToWindow
ipcr.Action = 1
'PrintResult% = ipcr.PrintReport
Else
'MsgBox "SERVICE AMOUNT IS NOT EQUAL TO TOTAL AMOUNT"
'PrintResult% = 20515
Exit Sub
End If

End Sub


Sub CHKINDOCNAME()
If Len(Trim(DBGrid1.Columns(4).Text)) <> 0 Then
       If rs_dnameflag = False Then
            RSDNAMEQR(0) = DBGrid1.Columns(4).Text
            Set rs_dname = RSDNAMEQR.OpenResultset(1, 3)
            rs_dnameflag = True
       Else
            RSDNAMEQR(0) = DBGrid1.Columns(4).Text
            rs_dname.MoreResults
            rs_dname.Requery
       End If
'Set rs_dname = rdb.OpenResultset("select * from ma102 where doc_code = " + DBGrid1.Columns(4).Text, 2, 3)
       If rs_dname.RowCount <> 0 Then
            DBGrid1.col = 5
            DBGrid1.Text = rs_dname("doc_name")
                        DBGrid1.row = DBGrid1.row + 1
            DBGrid1.col = 0
       Else
          If rs2("PAT_REG_TY") = "P" And rs2("U_D") = "D" Then
            MsgBox "NO CORRESPONDING DOCTOR NAME EXISTS....."
            DBGrid1.col = 4
            DBGrid1.SelStart = 0
            DBGrid1.SelLength = 5
            DBGrid1.SetFocus
        End If
       End If
 Else
  Call CHKDOCCODE
  
 End If
End Sub
'**1/7/99
Sub LVL_BTN_ENABL() '****12/6/99

Dim RS_LVL As rdoResultset
Dim RS_FM As rdoResultset

'Set RS_LVL = rdb.OpenResultset("select * from TASEC2 where TRIM(USR_GRP) = '" & Trim(USRGRP) & "' AND TRIM(SCREEN) = '" & UCase(Trim(reg_doc.Name)) & "'", 2, 3)
'Set RS_LVL = rdb.OpenResultset("select * from TOOLSADMN.scrn_auth1 where TRIM(USR_GRP) = '" & Trim(USRGRP) & "' AND TRIM(SCRN_ID) = '" & UCase(Trim(Me.Name)) & "'", 2, 3) 'COMMENTED BILLOPT 6/7/99
RS_LVLQR(0) = Trim(USRGRP)
RS_LVLQR(1) = UCase(Trim(Me.name))
Set RS_LVL = RS_LVLQR.OpenResultset(2, 3) 'BILLOPT 61299 ADDED


'**14/6/99
LVL_AUTH = 1
If RS_LVL.RowCount = 0 Then
MsgBox "USER NOT AUTHORISED TO ACCESS THE SCREEN"
LVL_AUTH = 0
Exit Sub
End If
'**14/6/99

If RS_LVL("LEVEL") = 4 Then
New1.Enabled = True
'update.Enabled = True
View.Enabled = True
Delete.Enabled = True
print1.Enabled = True
exit1.Enabled = True
Back.Enabled = True
ElseIf RS_LVL("LEVEL") = 3 Then
New1.Enabled = True
'update.Enabled = True
View.Enabled = True
Delete.Enabled = False
print1.Enabled = True
exit1.Enabled = True
Back.Enabled = True
ElseIf RS_LVL("LEVEL") = 2 Then
New1.Enabled = True
'update.Enabled = False
View.Enabled = True
Delete.Enabled = False
print1.Enabled = True
exit1.Enabled = True
Back.Enabled = True
ElseIf RS_LVL("LEVEL") = 1 Then
New1.Enabled = False
'update.Enabled = False
View.Enabled = True
Delete.Enabled = False
print1.Enabled = False
exit1.Enabled = True
Back.Enabled = True
ElseIf RS_LVL("LEVEL") = -1 Then
New1.Enabled = False
'update.Enabled = True
View.Enabled = True
Delete.Enabled = False
print1.Enabled = True
exit1.Enabled = True
Back.Enabled = True
End If

'***1/7/99


End Sub



Sub btn_vis(num)
        If num = 1 Then
            New1.Visible = False
            'Update.Visible = False
            View.Visible = False
            Delete.Visible = False
            exit1.Visible = False
            Back.Visible = True
            commit.Visible = True
            cancel.Visible = True
            print1.Visible = False
            
                       
        ElseIf num = 0 Then
            New1.Visible = False
           ' Update.Visible = True
            View.Visible = False
'            delete.Visible = True
            exit1.Visible = False
            Back.Visible = True
            commit.Visible = False
            cancel.Visible = False
           print1.Visible = True
        End If
        
End Sub



Sub del()
Dim i

If flag1 = 3 Then
i = MsgBox("do you want to delete ? ", vbYesNo)
If i = vbYes Then
 Set rs30 = rdb.OpenResultset("select * from ta107 where memo_no = '" & "M-" & M_number.Text & "'", rdOpenKeyset, 3)
 Set rs31 = rdb.OpenResultset("select * from ta1071 where memo_no = '" & "M-" & M_number.Text & "'", rdOpenKeyset, 3)
 Select Case rs30("AMT_RCVD")
Case "C"
     If ACCOUNTPSWD = "ACCOUNTS" Then
        MsgBox "PATIENT IS COMPANY REFFERED."
     Else
        MsgBox "PATIENT IS COMPANY REFFERED, YOU CAN'T DELETE"
        Call Refresh1
     Exit Sub
     End If
Case "T"
     If ACCOUNTPSWD = "ACCOUNTS" Then
        MsgBox "PATIENT IS TRUST AIDED."
     Else
        MsgBox "PATIENT IS TRUST AIDED, YOU CAN'T DELETE"
        Call Refresh1
        Exit Sub
     End If
Case "I"
    MsgBox "INTERIM BILL IS GENERATED FOR THIS MEMO, YOU CAN'T DELETE"
    Call Refresh1
    Exit Sub
Case "F"
    MsgBox "FINAL BILL IS GENERATED FOR THIS MEMO, YOU CAN'T DELETE"
    Call Refresh1
    Exit Sub
Case "R"
    MsgBox "RECEIPT IS GENERATED FOR THIS MEMO, YOU CAN'T DELETE"
    Call Refresh1
    Exit Sub
End Select
MsgBox "RECORD HAS BEEN DELETED ..."
 rs30.Delete
 While Not rs31.EOF
  rs31.Delete
  rs31.MoveNext
 Wend
 '**********************28499
 cleartxt
 M_number = " "
 lbl_cno = " "
 ReDim UserData(0 To 5, 0 To 20)
 DBGrid1.Refresh
 New1.SetFocus
'**********************28499

 Else
 MsgBox "RECORD NOT DELETED ..."
 Exit Sub
 End If
 End If
 flag1 = 0
End Sub

Sub new_view(cno As String)
On Error GoTo MSG
Dim CASENO As String
CASENO = Left(cno, 2) & "/" & Right(cno, 5)
'Set rs21 = rdb.OpenResultset("select ward_name,bed_no,FLOOR from MA104 where case_no = '" & CASENO & "'", rdOpenKeyset) 'BILLOPT 61299 COMMENTED
If rs21flag = False Then
    RS21QR(0) = CASENO
    Set rs21 = RS21QR.OpenResultset(1, 3)
    rs21flag = True
Else
    RS21QR(0) = CASENO
    rs21.MoreResults
    rs21.Requery
End If
'MsgBox rs21("ward_name")
'MsgBox rs21("floor")
'MsgBox rs21("bed_no")
'Set rs21 = rdb.OpenResultset("select ward_name,bed_type,bed_no,FLOOR from MA104 where case_no = '" & CASENO & "'", rdOpenKeyset) 'MOD 24599 selected bed_type
'Set rs18 = rdb.OpenResultset("select floor from MA104 where case_no = '" & caseno & "'", rdOpenKeyset)
If rs21.RowCount = 0 Then

Exit Sub
End If
lbl_wf_bed = RTrim(rs21("ward_name")) & " / " & RTrim(rs21("floor")) & "/ " & rs21("bed_no")
Exit Sub
MSG:  MsgBox Error(Err)
me_cno.SetFocus
End Sub

Sub ud_name()
'BILLOPT 61299 ADDED
d_u_name.clear
If lbl_cat = "C " Or lbl_cat = "NC" Then
    If rs14flag = False Then
        RS14QR(0) = "U"
        Set RS14 = RS14QR.OpenResultset(1, 3)
        rs14flag = True
    Else
        RS14QR(0) = "U"
        RS14.MoreResults
        RS14.Requery
    End If
Else
    If rs14flag = FLASE Then
        RS14QR(0) = "D"
        Set RS14 = RS14QR.OpenResultset(1, 3)
        rs14flag = True
    Else
        RS14QR(0) = "D"
        RS14.MoreResults
        RS14.Requery
    End If
End If
'RS14.Requery
RS14.MoveFirst
While Not RS14.EOF
  d_u_name.AddItem RS14("ud_name")
  RS14.MoveNext
Wend
End Sub
Sub dcode()
Dim opt As String

opt = lbl_cat
    
    Select Case opt
    
    Case "C ", "NC"
      If rs2("pat_reg_ty") = "R" Then '---121099
          Call doc_code
      End If '---121099
        
    'Case "A ", "B ", "D ", "FN"
    Case "A ", "B ", "D ", "F "
   
      Call doc_code
     ' SQL = "UPDATE TA107 SET DOC_CODE = 0 WHERE MEMO_NO = '" & MENO & "'"
      'rs4("doc_code") = rs33("doc_code")
    End Select
    
End Sub
Sub pat_type()
On Error GoTo MSG:
If rs2("pat_reg_ty") = "G" Then
   Label12.Caption = "UNIT NAME : "
   d_u_name.Visible = False
   txt_unit_doc_name.TabStop = False
   If flag = 1 Then
   txt_req_no.SetFocus
   End If
   'today
    lbl_dcode.Caption = "0"
       
ElseIf rs2("pat_reg_ty") = "P" Then
        Label12.Caption = "DOCTOR NAME : "
        If flag = 1 Then
        d_u_name.Visible = True
   End If
   If flag = 1 Then
   Call doc_code
   lbl_dcode.Caption = rs33("doc_code")
   Else
   lbl_dcode.Caption = rs22_CHK("DOC_CODE")
   End If
    
End If
Exit Sub
MSG:
 MsgBox "  DOCTOR CODE DOES NOT EXIST FOR  Dr.  " & rs2("ud_name")
 'txt_unit_doc_name
 
 btn_vis (0)
 cleartxt
 Exit Sub
 
End Sub

Sub patient_det(case_num As String)
'Dim d'gmk comm 29-4-00
'Dim d1'gmk comm 29-4-00
Dim pat_d As String
Dim pat_d1 As String
On Error GoTo MSG:
pat_det.Visible = True
invcat = False

'''HOME WORK

If flag = 2 Then
If rs22_CHKflag = False Then
    Exist_ChkQr(0) = M_number.Text '"M-" & M_number
    Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
    rs22_CHKflag = True
Else
    Exist_ChkQr(0) = M_number.Text
    rs22_CHK.MoreResults
    rs22_CHK.Requery
End If
'Set RSDNAMEQR = rdb.CreateQuery("IPRSDNAMEQ", "select * from ma102 where doc_code = ?") 'BILLOPT 61299 ADDED FOR DOCNAME IN DBGRID
 
 If rs_dnameflag = False Then
        
      RSDNAMEQR(0) = rs22_CHK("DOC_CODE")
      Set rs_dname = RSDNAMEQR.OpenResultset(1, 3)
      rs_dnameflag = True
 Else
    RSDNAMEQR(0) = rs22_CHK("DOC_CODE")
    rs_dname.MoreResults
    rs_dname.Requery
 End If
If rs_dname.RowCount = 0 Then
txt_unit_doc_name = "0"
Else
txt_unit_doc_name = Trim(rs_dname("doc_name"))
End If

Else
If rs2flag = False Then
    RS2QR(0) = case_num
    Set rs2 = RS2QR.OpenResultset(1, 3)
    rs2flag = True
Else
    RS2QR(0) = case_num 'BILLOPT 61299 ADDED
    rs2.MoreResults
    rs2.Requery
End If
'MsgBox rs2.RowCount
txt_unit_doc_name = Trim(rs2("UD_NAME"))
End If

''' HOME WORK
'This query is used to get the details of patients from Ta101
'RS2QR(0) = case_num 'BILLOPT 61299 ADDED
'Set rs2 = RS2QR.OpenResultset(2, 3) 'BILLOPT 61299 ADDED
If rs2flag = False Then
    RS2QR(0) = case_num
    Set rs2 = RS2QR.OpenResultset(1, 3)
    rs2flag = True
Else
    RS2QR(0) = case_num
    rs2.MoreResults
    rs2.Requery
End If
'MsgBox rs2.RowCount
If rs2("pat_reg_ty") = "P" Or rs2("pat_reg_ty") = "G" Then
If rs2("PAT_REG_TY") = "P" And rs2("U_D") = "D" Then
     DOCFLAG = True
   ' Call doc_code
ElseIf rs2("PAT_REG_TY") = "G" And rs2("U_D") = "U" Then
     DOCFLAG = False
Else
   MsgBox ("INVALID COMBINATION OF PAT REG AND U_D FLAG")
   invcat = True
   CANCEL_Click
   Exit Sub
End If
Else
  MsgBox ("INVALID PATIENT REG TYPE")
  invcat = True
  
  CANCEL_Click
  Exit Sub
  
End If

'''PARASOK
'today
If DOCFLAG = False And rs2("PAT_REG_TY") = "P" Then
           MsgBox "DOCTOR CODE PROBLEM WITH MEMO_NO"
           Exit Sub
End If

'today'''''Call doc_code '  gmk add on 12062000

Call pat_type

LBL_NAME = RTrim(rs2("first_name")) & " " & RTrim(rs2("mid_name")) & " " & rs2("last_name")
lbl_sex = rs2("sex")
If flag = 1 Then   '9/6/99 ADDED BECAUSE TO VIEW IP/OP MEMOS PATIENT CATG SHOULD BE THE CATEGORY AT THE TIME OF MEMO RAISING WHERE AS TA101 CATG IS LATEST CATG
   lbl_cat = rs2("pat_catg") '9/6/99 ADDED BECAUSE TO VIEW IP/OP MEMOS PATIENT CATG SHOULD BE THE CATEGORY AT THE TIME OF MEMO RAISING WHERE AS TA101 CATG IS LATEST CATG
End If '9/6/99 ADDED BECAUSE TO VIEW IP/OP MEMOS PATIENT CATG SHOULD BE THE CATEGORY AT THE TIME OF MEMO RAISING WHERE AS TA101 CATG IS LATEST CATG
'lbl_cat = rs2("pat_catg") '9/6/99 COMMENTED BECAUSE TO VIEW IP/OP MEMOS PATIENT CATG SHOULD BE THE CATEGORY AT THE TIME OF MEMO RAISING WHERE AS TA101 CATG IS LATEST CATG
lbl_typ = rs2("pat_type")

''''' ADDED BY SANDHYA DATED 6.07.2012 FOR BP PATIENT RAJIV GANDHI
If Trim(lbl_cat.Caption) = "BP" Then
  BPPATFLAG = True
End If
'''upto here


If rs2("PAT_TYPE") = "CR" And BPPATFLAG = False Then
If rs_crflag = False Then
    RS_CR_QR(0) = case_num
    Set rs_cr = RS_CR_QR.OpenResultset(1, 3)
    rs_crflag = True
Else
    RS_CR_QR(0) = case_num
    rs_cr.MoreResults
    rs_cr.Requery
End If
If rs_cr.RowCount <> 0 Then
'Set rs_cr = rdb.OpenResultset("select * from ta114 where case_no = '" & case_num & "'", rdOpenKeyset)
lbl_cr_card_no = rs_cr("cr_card")
Else
 MsgBox "Corresponding Credit Card No. does not exist...."
End If
ElseIf rs2("PAT_TYPE") = "TA" Then
If rs_trflag = False Then
    RS_TR_QR(0) = case_num
    Set rs_tr = RS_TR_QR.OpenResultset(1, 3)
    rs_trflag = True
Else
    RS_TR_QR(0) = case_num
    rs_tr.MoreResults
    rs_tr.Requery
 End If
If rs_tr.RowCount <> 0 Then
'Set rs_tr = rdb.OpenResultset("select * from ta113 where case_no = '" & case_num & "'", rdOpenKeyset)
lbl_tr_no = rs_tr("tr_acc_no")
Else
 MsgBox "Corresponding Trust Account No. does not exist...."
End If

'''' ADDED BY SANDHYA FOR SMART CARD PATIENT
ElseIf rs2("PAT_TYPE") = "CP" And scard = True Then

Dim RSTA105 As rdoResultset
Set RSTA105 = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.TA105 WHERE CASE_NO ='" & case_num & "' AND DIS_INT_DT IS NULL AND WARD_NAME IN (SELECT WARD_NAME FROM PATOOLSLIB.SMCWARD)", 1, 1)
'Set RSTA105 = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.TA105 WHERE CASE_NO ='" & case_num & "' AND trim(WARD_NAME) = 'HBB FIVE' AND DIS_INT_DT IS NULL ", 1, 1)
If RSTA105.RowCount = 1 Then
   Call GETSMARTCARD(case_num)
        lbl_smartcard.Visible = True
        lbl_deposit.Visible = True
        lbl_deposit.Caption = scard_bal_amt
        ''''ADDED FOR WRITING IN ERROR LOG
        
        ''''UPTO HERE
Else

lbl_smartcard.Visible = False
lbl_deposit.Visible = False
scard = False
End If

Else
lbl_cr_card_no = " "
lbl_tr_no = " "
End If
PATYPE = rs2("PAT_TYPE")
If PATYPE <> "CR" Then
If PATYPE <> "TA" Then
If PATYPE <> "CP" Then
    MsgBox "Incorrect Patient Type - " & PATYPE, vbExclamation, "Data Problem:TA101"
    Exit Sub
End If
End If
End If
pat_d = rs2("birth_date") 'gmk add 29-4-00
pat_d1 = Format(pat_d, "dd/mm/yyyy") 'gmk add 29-4-00
'lbl_age = Right(DateDiff("yyyy", me_date, d1), 2) 'gmk comm 29-4-00
lbl_age = Right(DateDiff("yyyy", me_date, pat_d1), 2) 'gmk add 29-4-00
'If age is less than zero or 0
'Then default value 1 is shown
If lbl_age <= 0 Then
    lbl_age = 1
End If

Exit Sub
MSG:  MsgBox Error(Err)
me_cno.SetFocus
me_cno.SelStart = 0
End Sub
Sub doc_code()
'This query gets the doctor code from DOCVIEW
If rs33flag = False Then
    RS33QR(0) = txt_unit_doc_name 'BILLOPT 61299 ADDED
    Set rs33 = RS33QR.OpenResultset(2, 3)
    rs33flag = True
 Else
    RS33QR(0) = txt_unit_doc_name
    rs33.MoreResults
    rs33.Requery
End If
'if no records are found..means the entry does not exist in DOCVIEW
 If ((rs33.RowCount) = 0) And (Trim(RS33QR(0)) <> "NO UNIT") Then 'GMD ADD FROM ON 21062000
     MsgBox "CODE DOES NOT EXIST FOR ...." & txt_unit_doc_name
     DOCFLAG = False
     
     Exit Sub
  Else
  'Indicates the the Doctor Code is there
      DOCFLAG = True
      If rs33.RowCount = 0 Then
        lbl_dcode = ""
      Else
       lbl_dcode = rs33("DOC_CODE")
      End If
 End If
End Sub
Sub assign()
On Error GoTo MSG:
Dim ass_d1 As String 'gmk add 29-4-00
Dim iRow As Integer
Dim opt As String
Dim bal_amt As Long
Dim trFlag As Boolean

bal_amt = 0
 
 
If AmtNotEq = 1 Then ''''2541999
    MsgBox " Either total is not tallying or service code missing, better you cancel and enter this memeo again. ", vbCritical
    AmtNotEq = 0
    Exit Sub
End If ''''2541999

'rdb.BeginTrans

If flag = 1 Then
  '******
    '17/10/2002-Sandhya
    'Patch Introduced To Find any Discrepancy In Memo Amount.
    'Before Saving.
       For iRow = 0 To mTotalRows&
        If Len(Trim(UserData(0, iRow))) > 0 Then
           If Not IsEmpty(UserData(0, iRow)) Then
               mser_amt = mser_amt + CDbl(UserData(3, iRow)) 'rs7("ser_amt")
           End If
        End If
       Next iRow
       
       If mser_amt <> CDbl(lbl_tot_amt) Then
          MsgBox "Total Memo Amount Does Not Match ! ", vbCritical
          Exit Sub
       End If


  '******
  'AUTO GENERATION OF MEMO NO.
me_mno.Mask = ">?#######"
S = Right(me_date.Text, 4)
If RS_SERIESflag = False Then
    RQ_SERIES(0) = "A" & S
    Set RS_SERIES = RQ_SERIES.OpenResultset(1, 3)
    RS_SERIESflag = True
Else
    RQ_SERIES(0) = "A" & S
    RS_SERIES.MoreResults
    RS_SERIES.Requery
End If

'*** start modified by v.s.k. on 13022012

If Not RS_SERIES.EOF Then
    K = Trim(RS_SERIES("SERIES")) + 1
Else
    K = 1
End If

trFlag = False
rdb.BeginTrans

Do
        rq_updtSERIES(0) = K
        rq_updtSERIES(1) = "A" & S
        rq_updtSERIES(2) = K
        rq_updtSERIES.Execute
        If rq_updtSERIES.RowsAffected > 0 Then
            trFlag = True
        Else
            K = K + 1
        End If
Loop While trFlag <> True
rdb.CommitTrans

AUTOMEMO = "M" & Right(RS_SERIES("YEAR1"), 1) & K
If Len(Trim(Mid(AUTOMEMO, 3))) <> 0 Then me_mno = padzero(6, "M", K, 1)
'*** end V.S.K.

'******** FOLLOWING CODE COMMENTED BY V.S.K. ON 13022012
'If RS_SERIES.RowCount > 0 Then
'AUTOMEMO = "M" & Right(RS_SERIES("YEAR1"), 1) & (Trim(RS_SERIES("SERIES")) + 1)
'
'Else
'   AUTOMEMO = "M" & Right(me_date.Text, 1) & "1"
'
'End If
''****
'If RS_SERIES.RowCount > 0 Then
'K = Mid(AUTOMEMO, 3, 8) ' ADDED 11/11/2001
'k1 = K
'If Len(Trim(Mid(AUTOMEMO, 3))) <> 0 Then
''    me_mno = pad(6, "M", me_mno, 1)
'
' If RS_SERIES.RowCount <> 0 Then   ' ADDED 11/11/2001
'    me_mno = padzero(6, "M", K, 1)  ' ADDED 11/11/2001
'  Else                              ' ADDED 11/11/2001
'      me_mno = padzero(6, "M", K, 1)
'  End If
'
'End If
'
''****
'        rq_updtSERIES(0) = k1
'        rq_updtSERIES(1) = "A" & S
'        rq_updtSERIES.Execute
' End If
'AUTO GENERATION OF MEMO NO

  
  
  
  
  
rdb.BeginTrans
  
  
  '****************************
  D1 = GET_REVDATE(me_date.Text)
   ''' SANDHYA PROTOCOL PATIENT STARTS
If Val(TXTNET) > 0 Then
    'If CDbl(Format(nettot, "########.00")) > 0 Then
   ''' PROTOCOL PATIENT ENDS
  
  addta107_QR(0) = me_mno.FormattedText
  addta107_QR(1) = "I"
  'addta107_QR(2) = Format(me_date.Text, "DD/MM/YYYY") gmk comm 29-4-00
  addta107_QR(2) = D1 'GET_REVDATE(me_date.Text) 'gmk add 29-4-00
  addta107_QR(3) = UCase(lbl_dno)
  addta107_QR(4) = CASENO
  'addta107_QR(5) = Trim(lbl_cat) ' comm on 24/5
  If Len(Trim(cbo_reqno.Text)) <> 0 Then 'add on 24/7
    addta107_QR(5) = Trim(lbl_cat) 'Trim(lbl_cat) 'rs4("MEMO_CATG") 'add on 24/7
  Else 'add on 24/7
    addta107_QR(5) = Trim(lbl_cat) 'add on 24/7
  End If 'add on 24/7
  
  If Len(Trim(txt_doc_sec_name)) = 0 Then
  addta107_QR(6) = "userid"
  Else
  addta107_QR(6) = txt_doc_sec_name
  End If
  addta107_QR(7) = ipadd & " - " & time1  'txt_rem
  'addta107_QR(8) = " " 'Trim(cbo_reqno.Text)  'txt_req_no ' comm on 25/7
  'addta107_QR(9) = Null ' comm on 25/7
'  If me_rdate.Text <> "" Then
'      ass_d1 = Left(me_rdate.Text, 2) & "/" & Mid(me_rdate.Text, 3, 2) & "/" & Right(me_rdate.Text, 4) 'Format(me_rdate.Text, "dd/mm/yyyy")
'    ass_d1 = GET_REVDATE(ass_d1)
'       addta107_QR(9) = ass_d1 ' gmk add 29-4-00
'  End If
'comm on 24/7
  If rs2("pat_reg_ty") = "G" Then
     addta107_QR(8) = 0
   Else
      'If ROWCNT = 0 Then
       '     addta107_QR(10) = 0 'rs33("DOC_CODE")
      'Else
            addta107_QR(8) = lbl_dcode '(rs33("DOC_CODE")) ' add on 25/7
      'End If
   End If
   '''' ADDED BY SANDHYA DATED 7.7.2012 FOR BP RGJAY PATIENT
   If rs2("pat_type") = "CR" And BPPATFLAG = False Then
     addta107_QR(9) = "C"
   '''' ADDED BY SANDHYA DATED 7.7.2012 FOR BP RGJAY PATIENT
   ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = True Then
        addta107_QR(9) = "R"
   ''' END OF BP PATIENT
   ElseIf rs2("pat_type") = "TA" Then
     addta107_QR(9) = "T"
   '''' ADDED BY SANDHYA DATED 14.7.2011 FOR SMART CARD
   
   ElseIf rs2("PAT_TYPE") = "CP" And scard = True Then
     addta107_QR(9) = "S"
    ''''UPTO HERE'''''
   Else
     addta107_QR(9) = Null
   End If
   
  '' PROTOCOL PATIENT STARTS
  If PROTFLG = True Then
     addta107_QR(10) = CDbl(Format(TXTNET, "#,##,##0.00"))
    Else
    addta107_QR(10) = CDbl(Format(lbl_tot_amt, "#,##,##0.00")) 'rs4("memo_amt")
  End If
  '''' PROTOCOL PATIENT ENDS
  addta107_QR(11) = "N"
  addta107_QR.Execute
  
  
End If
  
     
     
     For iRow = 0 To mTotalRows&
      If Len(Trim(UserData(0, iRow))) > 0 Then '''''28041999
        If Not IsEmpty(UserData(0, iRow)) Then
              '' SANDHYA PROTOCOL PATIENT STARTS
         If Val(TXTNET) > 0 Then

          PROTSERV_QR(0) = MIRBNO
          PROTSERV_QR(1) = Left(UserData(0, iRow), 4)
          Set PROTSERV_RS = PROTSERV_QR.OpenResultset(2, 3)
          PROTSERV_RS.Requery
          If PROTSERV_RS.RowCount = 0 Then
          ''' end of protocol
          
          
          addta1071_QR(0) = me_mno.FormattedText 'billopt 61299 added
          addta1071_QR(1) = UserData(0, iRow) 'billopt 61299 added
          addta1071_QR(2) = UserData(2, iRow) 'billopt 61299 added
          addta1071_QR(3) = Val(UserData(3, iRow)) 'billopt 61299 added
          
          If Val(UserData(4, iRow)) <> 0 Then 'billopt 61299 added
              addta1071_QR(4) = Val(UserData(4, iRow)) 'billopt 61299 added
          Else
              addta1071_QR(4) = 0
          End If 'billopt 61299 added
          
          If Len(Trim(UserData(0, iRow))) <> 0 Then 'add on 24/7
             If Len(Trim(UserData(0, iRow))) > 4 Then
              addta1071_QR(5) = Mid(UserData(0, iRow), 6, 13) 'ADD ON 24/7
              addta1071_QR(6) = GET_REVDATE(Mid(UserData(0, iRow), 20, 10)) ' add on 24/7
              addta1071_QR(7) = Mid(UserData(0, iRow), 31) ' add on 24/7
            Else  'add on 24/7
              addta1071_QR(5) = "" 'Mid(UserData(0, iRow), 6, 13) 'ADD ON 24/7
              addta1071_QR(6) = D1 ' CDate(Mid(UserData(0, iRow), 15)) ' add on 24/7
              '''' ADDED BY SANDHYA DATED 7.7.2012 FOR BP PATIENT RGJAYOJANA
              If BPPATFLAG = True Then
                 addta1071_QR(7) = "BP"
              Else
                 addta1071_QR(7) = "" ' add on 24/7
              End If
            End If  'add on 24/7
          Else
            addta1071_QR(5) = "" 'Mid(UserData(0, iRow), 6, 13) 'ADD ON 24/7
            addta1071_QR(6) = Null ' CDate(Mid(UserData(0, iRow), 15)) ' add on 24/7
            '''' ADDED BY SANDHYA DATED 7.7.2012 FOR BP PATIENT RGJAYOJANA
              If BPPATFLAG = True Then
                 addta1071_QR(7) = "BP"
              Else
                 addta1071_QR(7) = "" ' add on 24/7
              End If
          End If
          
          '*******
                    '17/10/2002-Parag/Sandhya
                    'Bifurcating Ser_amt as Prof_amt and Hosp_amt.
                    If ZRSGETSERVICEFLAG1 = False Then
                        zgetservice(0) = Trim(Left(UserData(0, iRow), 4))
                        Set zrsgetservice = zgetservice.OpenResultset(2, 1)
                        ZRSGETSERVICEFLAG1 = True
                    Else
                       zgetservice(0) = Trim(Left(UserData(0, iRow), 4))
                       zrsgetservice.MoreResults
                        zrsgetservice.Requery
                    End If
                    If zrsgetservice.RowCount > 0 Then
                        'Decide the break up amount
                        If Trim(Left(UserData(0, iRow), 4)) <> "Z" Then
                         Select Case Trim(lbl_cat)  ' Evaluate Number.
                         Case "A"
                        addta1071_QR(8) = UserData(2, iRow) * zrsgetservice!a_prof_ch
                        addta1071_QR(9) = UserData(2, iRow) * (zrsgetservice!a_hosp_ch - zrsgetservice!a_prof_ch)
                        addta1071_QR(10) = "X"
                         Case "B"
                        addta1071_QR(8) = UserData(2, iRow) * zrsgetservice!b_prof_ch
                        addta1071_QR(9) = UserData(2, iRow) * (zrsgetservice!b_hosp_ch - zrsgetservice!b_prof_ch)
                        addta1071_QR(10) = "X"
                         Case "F"
                        addta1071_QR(8) = UserData(2, iRow) * zrsgetservice!fn_prof_ch
                        addta1071_QR(9) = UserData(2, iRow) * (zrsgetservice!fn_hosp_ch - zrsgetservice!fn_prof_ch)
                        addta1071_QR(10) = "X"
                         Case "D"
                        addta1071_QR(8) = UserData(2, iRow) * zrsgetservice!d_prof_ch
                        addta1071_QR(9) = UserData(2, iRow) * (zrsgetservice!d_hosp_ch - zrsgetservice!d_prof_ch)
                        addta1071_QR(10) = "X"
                         Case "C"
                        addta1071_QR(8) = 0
                        addta1071_QR(9) = UserData(2, iRow) * zrsgetservice!c_ch
                        addta1071_QR(10) = "X"
                         Case "BP"
                        addta1071_QR(8) = 0
                        addta1071_QR(9) = UserData(2, iRow) * zrsgetservice!c_ch
                        addta1071_QR(10) = "X"
                         Case Else   ' "NC" case
                        addta1071_QR(8) = 0
                        addta1071_QR(9) = UserData(2, iRow) * zrsgetservice!nc_ch
                        addta1071_QR(10) = "X"
                         End Select
                        Else
                            addta1071_QR(8) = 0
                            addta1071_QR(9) = Val(UserData(3, iRow))
                            addta1071_QR(10) = "X"
                        End If
                        '''******add on 26/06/03
                        addta1071_QR(11) = CASENO
                End If
          
          
          '*******
          addta1071_QR.Execute 'billopt 61299 added
          
         End If
        End If
        '' PROTOCOL ENDS
          'added for ris on 11/05/2001
         If Len(Trim(UserData(0, iRow))) > 4 Then
             If RS_reqdet.RowCount <> 0 Then
                 rq_updtreqdet(0) = D1
                 rq_updtreqdet(1) = Trim(me_mno.FormattedText)
                 ' for icu req to by pass to write report directrily
                 
                 'made P from E for giving pathno in E-lab
                 If pathology.lbl_icureq.Visible = True Then
                   rq_updtreqdet(2) = "P"
                 Else
                   rq_updtreqdet(2) = "M"
                 End If
                 
                 rq_updtreqdet(3) = Trim(CASENO)
                 rq_updtreqdet(4) = Trim(Mid(UserData(0, iRow), 6, 13)) 'Trim(cbo_reqno.Text) 'txt_req_no.Text ' add on 25/7
                 'rq_updtreqdet(5) = lbl_dno
                ' rq_updtreqdet(6) = ""
                 rq_updtreqdet.Execute
             End If
            If Len(Mid(UserData(0, iRow), 6, 13)) <> 0 Then
             
             REQUISITN_UPDT(0) = "IPATIENT"
             REQUISITN_UPDT(1) = D1
             If pathology.lbl_icureq.Visible = True Then
                REQUISITN_UPDT(2) = "P" 'from E to P
             Else
                REQUISITN_UPDT(2) = "P"
             End If
             
             REQUISITN_UPDT(3) = Trim(Mid(UserData(0, iRow), 6, 13))
             REQUISITN_UPDT(4) = CASENO
             REQUISITN_UPDT.Execute
            End If
             
       If rs2("PAT_TYPE") = "TA" Then

            If Len(Mid(UserData(0, iRow), 6, 13)) <> 0 Then
            
              REQUISITN_UPDT(0) = "ITRUST"
              REQUISITN_UPDT(1) = me_date.FormattedText
              If pathology.lbl_icureq.Visible = True Then
                   REQUISITN_UPDT(2) = "P" ' from E to P for icu
              Else
                   REQUISITN_UPDT(2) = "P"
              End If
              
              REQUISITN_UPDT(3) = Mid(UserData(0, iRow), 6, 13)
              REQUISITN_UPDT(4) = CASENO
              REQUISITN_UPDT.Execute
                 
            End If
        End If
        
        If rs2("PAT_TYPE") = "CR" And BPPATFLAG = False Then

            If Len(Mid(UserData(0, iRow), 6, 13)) <> 0 Then
            
              REQUISITN_UPDT(0) = "ICOMPANY"
              REQUISITN_UPDT(1) = me_date.FormattedText
              If pathology.lbl_icureq.Visible = True Then
                     REQUISITN_UPDT(2) = "P"  ' from E to  P
              Else
                     REQUISITN_UPDT(2) = "P"
              End If
              
              REQUISITN_UPDT(3) = Mid(UserData(0, iRow), 6, 13)
              REQUISITN_UPDT(4) = CASENO
              REQUISITN_UPDT.Execute
                 
            End If
        End If
                         
         '''' ADDED BY SANDHYA FOR BP RAJIV GANDHI YOJANA
      If rs2("PAT_TYPE") = "CR" And BPPATFLAG = True Then
          If Len(Mid(UserData(0, iRow), 6, 13)) <> 0 Then
            REQUISITN_UPDT(0) = "RGJAY PAT"
            REQUISITN_UPDT(1) = me_date.FormattedText
            REQUISITN_UPDT(2) = "P"
            REQUISITN_UPDT(3) = Mid(UserData(0, iRow), 6, 13)
            REQUISITN_UPDT.Execute
          End If
      End If
                         
                         
       End If
         
        '''ADDED BY SANDHYA FOR CASHPAYING SMART CARD PATIENT
         If rs2("PAT_TYPE") = "CP" And scard = True Then
          If Len(Mid(UserData(0, iRow), 6, 13)) <> 0 Then
            REQUISITN_UPDT(0) = "ISMARTCARD"
            REQUISITN_UPDT(1) = me_date.FormattedText
            REQUISITN_UPDT(2) = "P"
            REQUISITN_UPDT(3) = Mid(UserData(0, iRow), 6, 13)
            REQUISITN_UPDT(4) = CASENO
            REQUISITN_UPDT.Execute
          End If
         End If
      '''' UPTO HERE ''''
              
         
           '' SANDHYA PROTOCOL PATIENT STARTS
   
      If PROTFLG = True Then
         If Len(Mid(UserData(0, iRow), 6, 13)) <> 0 Then
            
              REQUISITN_UPDT(0) = "IPROTOCOL"
              REQUISITN_UPDT(1) = me_date.FormattedText
              If pathology.lbl_icureq.Visible = True Then
                  REQUISITN_UPDT(2) = "P"  ' from E to  P
              End If
              REQUISITN_UPDT(3) = Mid(UserData(0, iRow), 6, 13)
              REQUISITN_UPDT(4) = CASENO
              REQUISITN_UPDT.Execute
                 
         End If
         
      End If
    ''' END OF PROTOCOL
 End If
End If ''28499R
Next iRow

If rs2("PAT_TYPE") = "TA" Then
If FLGTA = 1 Then
              Set rs_tr = rdb.OpenResultset("select * from PATOOLSLIB.ta113 where case_no = '" & CASENO & "'", 2, 3)
              rs_tr.Edit
              rs_tr("ser_amt") = CDbl(rs_tr("ser_amt")) + CDbl(lbl_tot_amt)
              rs_tr.Update

End If
ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = False Then
If FLGCR = 1 Then
              Set rs_cr = rdb.OpenResultset("select * from PATOOLSLIB.ta114 where case_no = '" & CASENO & "'", 2, 3)
              rs_cr.Edit
              rs_cr("ser_amt") = CDbl(rs_cr("ser_amt")) + CDbl(lbl_tot_amt)
              rs_cr.Update
End If

''''' ADDED BY SANDHYA DATED 7.7.2012 FOR BP RGJAY PATIENT
ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = True Then
     UPD_bppatn(0) = CDbl(lbl_tot_amt.Caption)
     UPD_bppatn(1) = CDbl(lbl_tot_amt.Caption)
     UPD_bppatn(2) = CASENO
     UPD_bppatn.Execute
     
     Set RS_BPPATN = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.BPPATIENT WHERE CASENO ='" & CASENO & "'", 1, 1)
     
     INSBPPATMEMO(0) = CASENO
     INSBPPATMEMO(1) = "BP"
     INSBPPATMEMO(2) = me_mno.FormattedText
     INSBPPATMEMO(3) = me_date.FormattedText
     INSBPPATMEMO(4) = CDbl(lbl_tot_amt)
     INSBPPATMEMO(5) = Trim(RS_BPPATN("SCHEM_CODE"))
     INSBPPATMEMO(6) = RS_BPPATN("SCHEM_AMT")
     INSBPPATMEMO(7) = usr_name
     INSBPPATMEMO(8) = Winsock1.LocalIP
     INSBPPATMEMO(9) = "I"
     INSBPPATMEMO.Execute
     
'''' END OF BP RGJAY PATIENT


'''' ADDED BY SANDHYA FOR SMART CARD PATIENT
 '''' TO UPDATE BALANCE OF SMART CARD PATIENT'S ACCOUNT
 
ElseIf rs2("PAT_TYPE") = "CP" And scard = True Then
If FLGSM = 1 Then
    Dim rdscard As rdoQuery
    Dim num As Integer
    Set rdscard = rdb.CreateQuery("", "update patoolslib.scardacc set ser_amt=ser_amt+?,bal_amt=bal_amt-? where case_no=? ")
    rdscard(0) = CDbl(lbl_tot_amt)
    rdscard(1) = CDbl(lbl_tot_amt)
    rdscard(2) = me_cno.FormattedText
    rdscard.Execute
    
    lbl_deposit.Caption = Val(lbl_deposit.Caption) - (lbl_tot_amt)
End If
End If
'''' END OF SMARTCARD

'' SANDHYA PROTOCOL PATIENT STARTS
If PROTFLG = True Then
  Set RSPROTCASE = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.PROTOCASE WHERE CASE_NO = '" & CASENO & "' AND IRBNO = '" & MIRBNO & "'", 2, 3)
  RSPROTCASE.Edit
  RSPROTCASE("SER_AMT") = CDbl(RSPROTCASE("SER_AMT")) + CDbl(TXTPRTTOT)
  RSPROTCASE.Update
End If
'****17/12/98

'' ADDED BY SANDHYA PROTOCOL PATIENT
'' TO INSERT INTO PROTOCOL MEMO TABLE
If Val(TXTPRTTOT) > 0 Then
For iRow = 0 To mTotalRows&
If Len(Trim(UserData(0, iRow))) > 0 Then
   If Not IsEmpty(UserData(0, iRow)) Then
   'PROTSERV_QR(0) = PROTCASE_RS("IRBno")
   PROTSERV_QR(0) = MIRBNO
   PROTSERV_QR(1) = Left(UserData(0, iRow), 4)
   Set PROTSERV_RS = PROTSERV_QR.OpenResultset(2, 3)
   PROTSERV_RS.Requery
 If PROTSERV_RS.RowCount = 1 Then
    PROTMEMO_Qr(0) = me_mno.FormattedText
    PROTMEMO_Qr(1) = me_date.FormattedText
    PROTMEMO_Qr(2) = Trim(PROTCASE_RS("IRBNO"))
    PROTMEMO_Qr(3) = CASENO
    PROTMEMO_Qr(4) = Left(UserData(0, iRow), 4)
    PROTMEMO_Qr(5) = Val(UserData(3, iRow))
    If zrsmemorcptflag = False Then
       zgetservice(0) = Trim(Left(UserData(0, iRow), 4))
       Set zrsgetservice = zgetservice.OpenResultset(2, 1)
       zrsmemorcptflag = True
    Else
       zgetservice(0) = Trim(Left(UserData(0, iRow), 4))
       zrsgetservice.MoreResults
       zrsgetservice.Requery
    End If
    If zrsgetservice.RowCount > 0 Then
                    'Decide the break up amount
      If Trim(Left(UserData(0, iRow), 4)) <> "Z" Then
       Select Case Trim(lbl_cat)  ' Evaluate Number.
       Case "A"
       PROTMEMO_Qr(6) = UserData(2, iRow) * zrsgetservice!a_prof_ch
       PROTMEMO_Qr(7) = UserData(2, iRow) * (zrsgetservice!a_hosp_ch - zrsgetservice!a_prof_ch)
       PROTMEMO_Qr(10) = "X"
       Case "B"
       PROTMEMO_Qr(6) = UserData(2, iRow) * zrsgetservice!b_prof_ch
       PROTMEMO_Qr(7) = UserData(2, iRow) * (zrsgetservice!b_hosp_ch - zrsgetservice!b_prof_ch)
       PROTMEMO_Qr(10) = "X"
       Case "F"
       PROTMEMO_Qr(6) = UserData(2, iRow) * zrsgetservice!fn_prof_ch
       PROTMEMO_Qr(7) = UserData(2, iRow) * (zrsgetservice!fn_hosp_ch - zrsgetservice!fn_prof_ch)
       PROTMEMO_Qr(10) = "X"
       Case "D"
       PROTMEMO_Qr(6) = UserData(2, iRow) * zrsgetservice!d_prof_ch
       PROTMEMO_Qr(7) = UserData(2, iRow) * (zrsgetservice!d_hosp_ch - zrsgetservice!d_prof_ch)
       PROTMEMO_Qr(10) = "X"
       Case "C"
       PROTMEMO_Qr(6) = 0
       PROTMEMO_Qr(7) = UserData(2, iRow) * zrsgetservice!c_ch
       PROTMEMO_Qr(10) = "X"
       Case Else   ' "NC" case
       PROTMEMO_Qr(6) = 0
       PROTMEMO_Qr(7) = UserData(2, iRow) * zrsgetservice!nc_ch
       PROTMEMO_Qr(10) = "X"
     End Select
   Else
       PROTMEMO_Qr(6) = 0
       PROTMEMO_Qr(7) = CDbl(UserData(3, iRow))
       PROTMEMO_Qr(10) = "X"
   
   End If
   PROTMEMO_Qr(8) = UserData(2, iRow)

PROTMEMO_Qr(9) = TXTPRTTOT
'PROTMEMO_Qr(7) = Mid(UserData(0, iRow), 6, 4)
If Val(UserData(4, iRow)) <> 0 Then
PROTMEMO_Qr(11) = Val(UserData(4, iRow)) 'rs7("doc_code")
Else
PROTMEMO_Qr(11) = 0  'rs7("doc_code")
End If
'PROTMEMO_Qr(8) = rs33("doc_code")
End If
PROTMEMO_Qr(12) = rs2("pat_type")
PROTMEMO_Qr(13) = Trim(rs2("pat_catg"))
PROTMEMO_Qr(14) = Mid(UserData(0, iRow), 6, 13)
PROTMEMO_Qr(15) = Mid(UserData(0, iRow), 20, 10)
PROTMEMO_Qr(16) = "I"
If rs2("PAT_TYPE") = "CR" Then
       PROTMEMO_Qr(17) = "C" 'rs4("amt_rcvd")
    ElseIf rs2("PAT_TYPE") = "TA" Then
        PROTMEMO_Qr(17) = "T"  'rs4("amt_rcvd")
    'V001 PATIENT DEPOSIT STARTS
    ElseIf rs2("PAT_TYPE") = "CP" And PATDEPO = True Then
        PROTMEMO_Qr(17) = "P"
    'V001 PATIENT DEPOSIT ENDS
    Else
        PROTMEMO_Qr(17) = " "
    End If


'PROTMEMO_Qr(17) = " "
If Len(Trim(txt_doc_sec_name)) = 0 Then
  PROTMEMO_Qr(18) = "userid"
  Else
PROTMEMO_Qr(18) = txt_doc_sec_name
  End If

PROTMEMO_Qr(19) = ipadd

'End If
      PROTMEMO_Qr.Execute
End If
End If
End If

Next iRow

End If



rdb.CommitTrans

rdb.BeginTrans
'' SANDHYA PROTOCOL PATIENT STARTS
If PROTFLG = False Then
    MAPING
Else
    IPROTMAP
End If
'' PROTOCOL PATIENT ENDS
End If
rdb.CommitTrans
'**** chandu
''    i = pathology.Data1.Recordset.RecordCount
''    'If MEMOFLAG = False And i <> 0 Then
''    '    MsgBox "Memo is not Generated"
''    '    btnenabl ("INI")
''    '    Exit Sub
''    'End If
''    Dim J As Integer
'' If i <> 0 Then
''
''    For J = 0 To i - 1
''    pathology.Data1.Recordset.MoveFirst
''    selreq_qr(0) = req_chglen(pathology.Data1.Recordset.Fields(0))
''    Set selreq_rs = selreq_qr.OpenResultset(1, 3)
''    If selreq_rs.RowCount > 0 Then
''
''            With pathology.Data1.Recordset
''                If Trim(selreq_rs("reqnstat")) = "M" Then
''                    .delete
''                    .MoveNext
''                End If
''
''                    If .EOF Then Exit Sub
''            End With
''    End If
''    Next J
''
''End If

'**** chandu

'''------------------------------------------------------
'''-------REQUISITION TABLE IS UPDATING IF THE PATIENT IS COMPANY REFERED OR TRUST ADIDED
''
''If rs2("PAT_TYPE") = "TA" Then
''RQ_reqdet1(0) = me_mno.FormattedText
''rs_reqdet1.MoreResults
''rs_reqdet1.Requery
''    If rs_reqdet1.RowCount > 0 Then
''        rs_reqdet1.MoveFirst
''       rdb.BeginTrans
''       While Not rs_reqdet1.EOF
''          REQUISITN_UPDT(0) = "TRUST"
''          REQUISITN_UPDT(1) = me_date.FormattedText
''          REQUISITN_UPDT(2) = "P"
''          REQUISITN_UPDT.Execute
''          rs_reqdet1.MoveNext
''        Wend
''        UPDTA107_QR(0) = "R"
''        UPDTA107_QR(1) = me_mno.FormattedText
''        UPDTA107_QR.Execute
''        rdb.CommitTrans
''    End If
''
''
''End If
''
''If rs2("PAT_TYPE") = "CR" Then
''
''RQ_reqdet1(0) = me_mno.FormattedText
''rs_reqdet1.MoreResults
''rs_reqdet1.Requery
''    If rs_reqdet1.RowCount > 0 Then
''        rs_reqdet1.MoveFirst
''        rdb.BeginTrans
''        While Not rs_reqdet1.EOF
''          REQUISITN_UPDT(0) = "COMPANY"
''          REQUISITN_UPDT(1) = me_date.FormattedText
''          REQUISITN_UPDT(2) = "P"
''          REQUISITN_UPDT.Execute
''          rs_reqdet1.MoveNext
''        Wend
''        UPDTA107_QR(0) = "R"
''        UPDTA107_QR(1) = me_mno.FormattedText
''        UPDTA107_QR.Execute
''
''        rdb.CommitTrans
''    End If
''
''End If
'''-------------



'------------------------------------------------------

'*********************************
     
     
     
     
   
'  addta107_QR(0) = "M-" & me_mno.Text
'  addta107_QR(1) = "I"
'  'addta107_QR(2) = Format(me_date.Text, "DD/MM/YYYY") gmk comm 29-4-00
'  addta107_QR(2) = GET_REVDATE(me_date.Text) 'gmk add 29-4-00
'  addta107_QR(3) = UCase(lbl_dno)
'  addta107_QR(4) = CASENO
'  addta107_QR(5) = lbl_cat
'  addta107_QR(6) = txt_doc_sec_name
'  addta107_QR(7) = txt_rem
'  addta107_QR(8) = txt_req_no
'  addta107_QR(9) = Null
'  If me_rdate.Text <> "" Then
'   ' d1 = Left(me_rdate.Text, 2) & "/" & Mid(me_rdate.Text, 3, 2) & "/" & Right(me_rdate.Text, 4) 'Format(me_rdate.Text, "dd/mm/yyyy")
'   'd1 = Format(d1, "YYYY-MM-DD")
'    ass_d1 = Left(me_rdate.Text, 2) & "/" & Mid(me_rdate.Text, 3, 2) & "/" & Right(me_rdate.Text, 4) 'Format(me_rdate.Text, "dd/mm/yyyy")
'    ass_d1 = GET_REVDATE(ass_d1)
'    'addta107_QR(9) = d1 ' gmk comm 29-4-00
'    addta107_QR(9) = ass_d1 ' gmk add 29-4-00
'  End If
'  If rs2("pat_reg_ty") = "G" Then
'     addta107_QR(10) = 0
'   Else
'      'If ROWCNT = 0 Then
'       '     addta107_QR(10) = 0 'rs33("DOC_CODE")
'      'Else
'            addta107_QR(10) = (rs33("DOC_CODE"))
'      'End If
'   End If
'   If rs2("pat_type") = "CR" Then
'     addta107_QR(11) = "C"
'   ElseIf rs2("pat_type") = "TA" Then
'     addta107_QR(11) = "T"
'   Else
'     addta107_QR(11) = Null
'   End If
'
'  addta107_QR(12) = Format(lbl_tot_amt, "#,##,##,##,##0.00")
'
'  addta107_QR.Execute
  
  
  
  
  
  
  
  
  
  
  
  
  'SQL = "INSERT INTO TA107(MEMO_NO,IO_MEMO,MEMO_DATE,DEPT_CODE,CASE_NO,MEMO_CATG,D_S_NAME,REM,REQ_NO) VALUES ('" & MENO & "','I','" & MEDT & "','" & UCase(lbl_dno) & "','" & CASENO & "','" & lbl_cat & "','" & txt_doc_sec_name & "','" & txt_rem & "'," & Val(txt_req_no) & ")"
  
  
  
  
  
 ' SQL = "INSERT INTO TA107(MEMO_NO) VALUES ('" & MENO & "')"

  'rdb.Execute SQL
  
   'rs4("memo_no") = "M-" & me_mno.Text
 ' End If
'    rs4("io_memo") = "I"
'    rs4("memo_date") = Format(me_date.Text, "dd/mm/yyyy")
'    'rs4("dept_code") = UCase(Com_dno.Text)
'    rs4("dept_code") = lbl_dno ''19599
'    rs4("case_no") = CASENO 'me_cno
'     rs4("MEMO_CATG") = lbl_cat '9/6/99 ADDED BECAUSE TO VIEW IP/OP MEMOS PATIENT CATG SHOULD BE THE CATEGORY AT THE TIME OF MEMO RAISING WHERE AS TA101 CATG IS LATEST CATG , IN MUMBAI IT SHOULD BE "MEMO_CATG" REMOVE COMMENT FOR bills
'    rs4("d_s_name") = txt_doc_sec_name
'    rs4("rem") = txt_rem
'    rs4("req_no") = txt_req_no

'    If me_rdate.Text <> "" Then
'       rs4("req_date") = Left(me_rdate.Text, 2) & "/" & Mid(me_rdate.Text, 3, 2) & "/" & Right(me_rdate.Text, 4) 'Format(me_rdate.Text, "dd/mm/yyyy")
'    End If
   
 'Dim REQDT
'  REQDT = Left(me_rdate.Text, 2) & "/" & Mid(me_rdate.Text, 3, 2) & "/" & Right(me_rdate.Text, 4) 'Format(me_rdate.Text, "dd/mm/yyyy")
  'REQDT = Right(me_rdate.Text, 4) & "-" & Mid(me_rdate.Text, 3, 2) & "-" & Left(me_rdate.Text, 2)  'Format(me_rdate.Text, "dd/mm/yyyy")

    'If me_rdate.Text <> "" Then
      'SQL = "UPDATE TA107 SET REQ_DATE = '" & REQDT & "' WHERE MEMO_NO = '" & MENO & "'"
      'rdb.Execute SQL
    'End If
    
   
    'Call dcode
   'If rs2("pat_reg_ty") = "G" Then
   'SQL = "UPDATE TA107 SET DOC_CODE = 0 WHERE MEMO_NO = '" & MENO & "'"
   'Else
   'SQL = "UPDATE TA107 SET DOC_CODE = " & rs33("DOC_CODE") & " WHERE MEMO_NO = '" & MENO & "'"
   'End If
   
     'rdb.Execute SQL

    
    
'    If rs2("pat_reg_ty") = "G" Then
'    rs4("doc_code") = Val(" ")
'    Else
'    rs4("doc_code") = rs33("doc_code")
'    End If
    
'   If rs2("pat_type") = "CR" Then
'   SQL = "UPDATE TA107 SET AMT_RCVD = 'C' WHERE MEMO_NO = '" & MENO & "'"
'   ElseIf rs2("pat_type") = "TA" Then
'   SQL = "UPDATE TA107 SET AMT_RCVD = 'T' WHERE MEMO_NO = '" & MENO & "'"
'   End If
'    rdb.Execute SQL
    
    
'    If rs2("PAT_TYPE") = "CR" Then
'    rs4("amt_rcvd") = "C"
'    ElseIf rs2("PAT_TYPE") = "TA" Then
'    rs4("amt_rcvd") = "T"
'    End If
    
  'Dim SERCD As String
  
    
'For iRow = 0 To mTotalRows&
'      If Len(Trim(UserData(0, iRow))) > 0 Then '''''28041999
'            If Not IsEmpty(UserData(0, iRow)) Then
'
'
'  '*****
'  'SQL = "INSERT INTO TA1071(MEMO_NO,SER_CODE,QTY,SER_AMT) VALUES ('" & MENO & "','" & UserData(0, iRow) & "'," & UserData(2, iRow) & "," & Val(UserData(3, iRow)) & ")" 'billopt 61299 commented
'  'rdb.Execute SQL 'billopt 61299 commented
'  '*****
'
' ' SERCD = UserData(0, iRow)
'
'          addta1071_QR(0) = "M-" & me_mno.Text 'billopt 61299 added
'          addta1071_QR(1) = UserData(0, iRow) 'billopt 61299 added
'          addta1071_QR(2) = UserData(2, iRow) 'billopt 61299 added
'          addta1071_QR(3) = Val(UserData(3, iRow)) 'billopt 61299 added
'
'          If Val(UserData(4, iRow)) <> 0 Then 'billopt 61299 added
'              addta1071_QR(4) = Val(UserData(4, iRow)) 'billopt 61299 added
'          Else
'              addta1071_QR(4) = 0
'          End If 'billopt 61299 added
'
'          addta1071_QR.Execute 'billopt 61299 added
'
'
'
''                rs7("memo_no") = "M-" & me_mno.Text
''                rs7("ser_code") = UserData(0, iRow)
''                rs7("qty") = UserData(2, iRow)
''                rs7("ser_amt") = Val(UserData(3, iRow))
''billopt 61299 commented
'
''If Val(UserData(4, iRow)) <> 0 Then
''   SQL = "UPDATE TA1071 SET DOC_CODE = " & Val(UserData(4, iRow)) & " WHERE MEMO_NO = '" & MENO & "' AND SER_CODE = '" & UserData(0, iRow) & "'"
''   Else
''   SQL = "UPDATE TA1071 SET DOC_CODE = 0 WHERE MEMO_NO = '" & MENO & "'AND SER_CODE = '" & UserData(0, iRow) & "'"
''   End If
'' rdb.Execute SQL
'''billopt 61299 commented
''                If Val(UserData(4, iRow)) <> 0 Then
''                'If Not DBGrid1.Columns(4) = " " Then
''                    rs7("doc_code") = Val(UserData(4, iRow))
''                Else
''                    rs7("doc_code") = 0 'lbl_dcode '***8/12/98
''                End If
''
''                rs7.update
''                rs7.AddNew
'
'            End If
'        End If ''28499R
'Next iRow

' rs7.update
'rs4("memo_amt") = Format(lbl_tot_amt, "#,##,##0.00")
'            rs4.update
           
 'billopt 61299 commented
           
           
'          Dim TOTMEAMT As Long
' TOTMEAMT = Format(lbl_tot_amt, "#,##,##0.00")
'
'    SQL = "UPDATE TA107 SET MEMO_AMT = " & TOTMEAMT & " WHERE MEMO_NO = '" & MENO & "'"
'  ' End If
'    rdb.Execute SQL
  
'billopt commented
           

'If rs2("PAT_TYPE") = "TA" Then
'If FLGTA = 1 Then
'              rs_tr.Edit
'              rs_tr("ser_amt") = rs_tr("ser_amt") + lbl_tot_amt
'              rs_tr.Update
'
'End If
'ElseIf rs2("PAT_TYPE") = "CR" Then
'If FLGCR = 1 Then
'              'Set rs_cr = rdb.OpenResultset("select * from ta114 where case_no = '" & CASENO & "'", 2, 3) Dim addta107_QR As rdoQuery 'BILLOPT 61299 commented
'              rs_cr.Edit
'              rs_cr("ser_amt") = rs_cr("ser_amt") + lbl_tot_amt
'              rs_cr.Update
'End If
'End If
''****17/12/98
'rdb.CommitTrans

''ADDED BY MONO
'MsgBox " Please Take Print Out to Complete the Transaction", vbInformation, "MEMO"
Exit Sub
MSG:
rdb.RollbackTrans
 MsgBox Error(Err)
'Resume Next
Exit Sub
End Sub
Sub add_edit()
If AmtNotEq = 1 Then '19599
    MsgBox " Either total is not tallying or service code missing ", vbCritical
    AmtNotEq = 0
    Exit Sub
End If '19599



assign
End Sub

Private Sub Back_Click()
'    Dim r1 As rdoResultset
'    For Each r1 In rdb.rdoResultsets
'
'    r1.Close
'    Next
REQGOTFLAG = False
ReDim UserData(0 To 5, 0 To 12)
DBGrid1.Refresh
cleartxt
Me.Hide

'MENU1.Show
'tmh_ris_mnu.Show   '**30/6/99
'SELMNU.Show
End Sub
Private Sub TW()
REQNO = Left(cbo_reqno.Text, 3) & Mid(cbo_reqno.Text, 5, 2) & Mid(cbo_reqno.Text, 8, 2) & Right(cbo_reqno.Text, 6)
    If RS_reqdetflag = False Then
        rq_reqdet(0) = CASENO
        rq_reqdet(1) = "IP"
        rq_reqdet(2) = REQNO
        'rq_reqdet(3) = lbl_dno
        'rq_reqdet(4) = ""
        Set RS_reqdet = rq_reqdet.OpenResultset(1, 1)
        RS_reqdetflag = True
    Else
        rq_reqdet(0) = CASENO
        rq_reqdet(1) = "IP"
        rq_reqdet(2) = REQNO
        'rq_reqdet(3) = lbl_dno
        'rq_reqdet(4) = ""
        RS_reqdet.MoreResults
        RS_reqdet.Requery
        
    End If
End Sub


Private Sub cbo_reqno_GotFocus()
If REQGOTFLAG = False Then
REQGOTFLAG = True
AmtNotEq = 0
CASENO = pathology.msk_caseno.FormattedText

'''' ADDED BY SANDHYA FOR SMART CARD PATIENT
If Trim(pathology.PATYPE) = "CP" Then

Dim RSTA105 As rdoResultset
Set RSTA105 = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.TA105 WHERE CASE_NO ='" & pathology.msk_caseno.FormattedText & "' AND DIS_INT_DT IS NULL AND WARD_NAME IN (SELECT WARD_NAME FROM PATOOLSLIB.SMCWARD)", 1, 1)
'Set RSTA105 = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.TA105 WHERE CASE_NO ='" & pathology.msk_caseno.FormattedText & "' AND trim(WARD_NAME) = 'HBB FIVE' AND DIS_INT_DT IS NULL ", 1, 1)
If RSTA105.RowCount = 1 Then
   scard = True
   
   Call GETSMARTCARD(CASENO)
        lbl_smartcard.Visible = True
        lbl_deposit.Visible = True
        lbl_deposit.Caption = scard_bal_amt
        ''''ADDED FOR WRITING IN ERROR LOG
        
        ''''UPTO HERE
Else

lbl_smartcard.Visible = False
lbl_deposit.Visible = False
scard = False
End If
End If
i = pathology.Data1.Recordset.RecordCount
  If i <> 0 Then
    pathology.Data1.Recordset.MoveFirst
    While Not pathology.Data1.Recordset.EOF
        '
     '****  filling the reqno combo box
   
        
        If RS_reqdetflag = False Then
            rq_reqdet(0) = CASENO
            rq_reqdet(1) = "IP"
            rq_reqdet(2) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
            'rq_reqdet(3) = lbl_dno
            'rq_reqdet(4) = ""
            Set RS_reqdet = rq_reqdet.OpenResultset(1, 3)
            RS_reqdetflag = True
        Else
            rq_reqdet(0) = CASENO
            rq_reqdet(1) = "IP"
            rq_reqdet(2) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
            'rq_reqdet(3) = lbl_dno
            'rq_reqdet(4) = ""
            RS_reqdet.MoreResults
            RS_reqdet.Requery
        End If
        If RS_reqdet.RowCount > 0 Then
                If Trim(RS_reqdet("ser_code")) = "8888" Then
                    ''***ADD BY MONO FOR BP CROSSMATCH CHARGE*************
                    dept = Mid(req_chglen(pathology.Data1.Recordset.Fields("reqno")), 4, 2)
                    If dept = "BP" Then
                        If SELH002_FLAG = False Then
                            SELH002_QR(0) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                            Set SELH002_RS = SELH002_QR.OpenResultset(1, 3)
                            SELH002_FLAG = True
                        Else
                            SELH002_QR(0) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                            SELH002_RS.MoreResults
                            SELH002_RS.Requery
                        End If
                        If SELH002_RS.RowCount > 0 Then
                            While Not SELH002_RS.EOF
                                cbo_reqno.AddItem SELH002_RS("PRO_code") & " " & req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                                SELH002_RS.MoveNext
                            Wend
                        End If
                    ''***ADD BY MONO FOR BP CROSSMATCH CHARGE*************
                    ElseIf dept = "MP" Then
                        If SELDMOLEPROCFLAG = False Then
                            SELDMOLEPROCQR(0) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                            Set SELDMOLEPROCRS = SELDMOLEPROCQR.OpenResultset(1, 3)
                            SELDMOLEPROCFLAG = True
                        Else
                            SELDMOLEPROCQR(0) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                            SELDMOLEPROCRS.MoreResults
                            SELDMOLEPROCRS.Requery
                        End If
                        If SELDMOLEPROCRS.RowCount > 0 Then
                            While Not SELDMOLEPROCRS.EOF
                                cbo_reqno.AddItem SELDMOLEPROCRS("PRO_code") & " " & req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                                SELDMOLEPROCRS.MoveNext
                            Wend
                        End If
                    Else
                        If SELDMULPROCEFLAG = False Then
                            SELDMULPROCEQR(0) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                            Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
                            SELDMULPROCEFLAG = True
                        Else
                            SELDMULPROCEQR(0) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                            SELDMULPROCERS.MoreResults
                            SELDMULPROCERS.Requery
                        End If
                        If SELDMULPROCERS.RowCount > 0 Then
                            While Not SELDMULPROCERS.EOF
                                cbo_reqno.AddItem SELDMULPROCERS("PRO_code") & " " & req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                                SELDMULPROCERS.MoveNext
                            Wend
                        End If
                    End If
                Else
                'ADDED BY SHWETA FOR FROZEN SECTION MEMO 'F309 10-03-2013
                    If Left(req_chglen(pathology.Data1.Recordset.Fields("reqno")), 5) = "FZZSP" Then
                    
                    If SELDMULPROCEFLAG = False Then
                            SELDMULPROCEQR(0) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                            Set SELDMULPROCERS = SELDMULPROCEQR.OpenResultset(1, 3)
                            SELDMULPROCEFLAG = True
                        Else
                            SELDMULPROCEQR(0) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                            SELDMULPROCERS.MoreResults
                            SELDMULPROCERS.Requery
                        End If
                        
                        If SELDMULPROCERS.RowCount > 0 Then
                            While Not SELDMULPROCERS.EOF
                                cbo_reqno.AddItem SELDMULPROCERS("PRO_code") & " " & req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                                SELDMULPROCERS.MoveNext
                            Wend
                        End If
                    
                    Else
                
                    cbo_reqno.AddItem RS_reqdet("ser_code") & " " & req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                    End If
                End If
        End If
       
 
    '**** filling the reqno combo box
    
    
        'cbo_reqno.AddItem pathology.Data1.Recordset.Fields("reqno")
              
    
       pathology.Data1.Recordset.MoveNext
    Wend
 ' becuase reqno is ois becoming blank
 
 pathology.Data1.Recordset.MoveFirst
 End If
 ReDim UserData(0 To 5, 0 To 20)
 DBGrid1.Refresh
 End If
End Sub

Private Sub cbo_reqno_LostFocus()
'If Len(Trim(cbo_reqno.Text)) = 0 Then
''txt_req_no.Text = "  "
'cbo_reqno.ListIndex = -1
''added for ris modification on 11052001
'
'TW
'
'Else
'''' ADDED BY SHWETA DATED 10.1.2012 FOR SKIPING GENERAL CASE TO CHECK DOCCODE
 If Trim(lbl_cat.Caption) <> "C" And Trim(lbl_cat.Caption) <> "NC" Then
 doc_code
 End If
 
 CASENO = me_cno.FormattedText
 If rs2flag = False Then
    RS2QR(0) = CASENO
    Set rs2 = RS2QR.OpenResultset(1, 3)
    rs2flag = True
 Else
    RS2QR(0) = CASENO
    rs2.MoreResults
    rs2.Requery
 End If
 'REMOVED COMMENT BY AMARJEET ON 02-11-2002
 'TW
 'TW
    
  If RS_reqdet.RowCount = 0 Then
    MsgBox "no corresponding req. no. record exists in rrequisitn table"
  Else
  
   
    'me_rdate = rs_reqdet("reqdate")
    ReDim Preserve UserData(0 To 5, 0 To 12)
    'DBGrid1.Refresh
    DBGrid1.col = 0
    'DBGrid1.row = 0
    
    'DBGrid1.Text = RS_reqdet("ser_code") & " " & reqno & " " & RS_reqdet("reqndt") & " " & Trim(RS_reqdet("pat_catg"))
    If Trim(cbo_reqno.Text) <> "" Then
        DBGrid1.Text = cbo_reqno & " " & RS_reqdet("reqndt") & " " & Trim(RS_reqdet("pat_catg"))
    End If
    me_rdate.Text = RS_reqdet("reqndt")
    flag = 1
    If flag = 1 Then
       lbl_cat = RS_reqdet("pat_catg")
    End If
    
    'DBGrid1.Col = 5
    'DBGrid1.Text = rs_reqdet("DOCREF")
    
   End If
'added for ris modification on 11052001'
  
'End If

'ADD ON 24/7
If Len(Trim(DBGrid1.Columns(0).Text)) = 0 Then
          DBGrid1.col = 0
          DBGrid1.Enabled = True
          DBGrid1.SetFocus
          DBGrid1.SelStart = 0
Else
          DBGrid1.col = 0
          DBGrid1.Enabled = True
          DBGrid1.SetFocus
          '
          DBGrid1.SelText = DBGrid1.Columns(0).Text
          DBGrid1.SelStart = 4
End If
          
          DBGrid1.col = 0
          DBGrid1.Enabled = True
          DBGrid1.SetFocus
          
If cbo_reqno.ListCount <> 0 Then
    If Trim(cbo_reqno.Text) <> "" Then
        cbo_reqno.RemoveItem cbo_reqno.ListIndex
    End If
End If
'ADD ON 24/7
End Sub

Private Sub Com_dno_Click()
On Error Resume Next
'Set rs2 = rdb.OpenResultset("select dept_code from MA106 where dept_name = '" & UCase(Com_dno.Text) & "'", rdOpenKeyset) 'commented ON 13/8/99
'lbl_dno = rs2("dept_code") 'COMMENTED ON 9/7/99 SETUP

'Set RS19 = rdb.OpenResultset("select dept_code from MA106 where dept_name = '" & UCase(Com_dno.Text) & "'", rdOpenKeyset) 'ADDED ON 13/8/99
'rs12.Move First
'rs12.Requery
'While Not rs12.EOF
'If Trim(rs12("DEPT_NAME")) = Trim(Com_dno.Text) Then
'    lbl_dno = rs12("dept_code")
'    rs12.MoveLast
'End If
'    rs12.MoveNext
'Wend
'me_cno.SetFocus

End Sub

'Private Sub Com_cno_Click()
'me_cno = Com_cno.Text
'me_cno.SetFocus
'me_cno.SelStart = 8

'd_u_name.Clear
'Call patient_det(me_cno)
'Call ud_name
'new_view (Com_cno)
'End Sub



Private Sub Com_dno_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then

''Set rs2 = rdb.OpenResultset("select dept_name from MA106 where dept_code = '" & UCase(Com_dno.Text) & "'", rdOpenKeyset)
'Set rs2 = rdb.OpenResultset("select dept_code from MA106 where dept_name = '" & UCase(Com_dno.Text) & "'", rdOpenKeyset)
''lbl_dname = rs2("dept_name")
'lbl_dno = rs2("dept_code")
''com_cno.SetFocus
'me_cno.SetFocus
Com_dno_LostFocus
End If
End Sub


Private Sub Com_dno_LostFocus() 'changed on 19599
On Error Resume Next
''Set rs2 = rdb.OpenResultset("select dept_code from MA106 where dept_name = '" & UCase(Com_dno.Text) & "'", rdOpenKeyset) 'commented ON 13/8/99
''lbl_dno = rs2("dept_code") 'COMMENTED ON 9/7/99 SETUP
'
''Set RS19 = rdb.OpenResultset("select dept_code from MA106 where dept_name = '" & UCase(Com_dno.Text) & "'", rdOpenKeyset) 'ADDED ON 13/8/99
''rs12.Move First
rs12.Requery
While Not rs12.EOF
If Trim(rs12("DEPT_NAME")) = Trim(Com_dno.Text) Then
    lbl_dno = rs12("dept_code")
    rs12.MoveLast
End If
    rs12.MoveNext
Wend

End Sub

Private Sub commit_GotFocus() '19599
VALIDATEAMT
End Sub

Private Sub d_u_name_Click()
If Len(Trim(d_u_name)) = 0 Then
 Exit Sub
 End If
txt_unit_doc_name.Text = d_u_name
Call doc_code
If rs2("pat_reg_ty") = "G" Then
lbl_dcode = " "
Else
lbl_dcode = rs33("doc_code")
End If
txt_req_no.SetFocus
End Sub


Private Sub DBGrid1_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo MSG

Dim opt As String
Dim ncchg As String ' commented for z-series memos***4899

If ColIndex = 0 Then
  DBgrid1_KeyPress (13)
  If NOCUR = 1 Then
   Exit Sub
  End If
  If ncchg = "Z" Then  ' added for z-series memos***4899

    DBGrid1.Columns(3).Locked = False 'ADDED ECTV
    DBGrid1.Columns(4).Locked = True ''ADDED ECTV
    DBGrid1.Columns(5).Locked = True 'ADDED ECTV
    flex2.Enabled = False
    DBGrid1.col = 2
   
  ElseIf ncchg <> "Z" And RS18.RowCount <> 0 Then
    DBGrid1.Columns(4).Locked = False ''ADDED ECTV
    DBGrid1.Columns(5).Locked = False 'ADDED ECTV
    flex2.Enabled = True
    'DBGrid1.Columns(3).Locked = True
    
    'DBGrid1.Col = 2
  ElseIf ncchg <> "Z" And RS18.RowCount = 0 Then
    DBGrid1.col = 0
    DBGrid1.SelStart = 0 ''21899
    DBGrid1.SetFocus
    Exit Sub
  End If ' added for z-series memos***4899

  
ElseIf ColIndex = 2 Then
  
  If rs18flag = False Then
    RS18QR(0) = DBGrid1.Columns(0)
    Set RS18 = RS18QR.OpenResultset(1, 3)
    rs18flag = True
  Else
    RS18QR(0) = DBGrid1.Columns(0)
    RS18.MoreResults
    RS18.Requery
  End If
  'Set RS18 = rdb.OpenResultset("select * from MA101 where ser_code = '" & DBGrid1.Columns(0) & "'", rdOpenKeyset, 3) 'BILLOPT 61299 COMMENTED

'opt = lbl_cat 'comented on 24599
ncchg = Left(DBGrid1.Columns(0), 1) '***22/10/98
''24599

If rsdbtpflag = False Then
    RSDBTPQR(0) = Trim(rs21("bed_type"))
    RSDBTPQR(1) = Trim(lbl_cat)
    Set RSDBTP = RSDBTPQR.OpenResultset(2, 3)
    rsdbtpflag = True
Else
    RSDBTPQR(0) = Trim(rs21("bed_type"))
    RSDBTPQR(1) = Trim(lbl_cat)
    RSDBTP.MoreResults
    RSDBTP.Requery
End If
'MsgBox RSDBTP.RowCount
'Set RSDBTP = rdb.OpenResultset("select ch_code from ma107 where pat_cat = '" & lbl_cat & "' and bed_type = '" & rs21("bed_type") & "'", 2, 3) 'BILLOPT 61299 COMMENTED

opt = Trim(RSDBTP("ch_code"))
''24599
Select Case opt
Case "C"
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH")
Case "B"
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("B_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("B_HOSP_CH") - RS18("B_PROF_CH"))  '***22/10/98
    End If
        
Case "A"
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("A_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("A_HOSP_CH") - RS18("A_PROF_CH"))  '***22/10/98
    End If
    
Case "D"
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("D_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("D_HOSP_CH") - RS18("D_PROF_CH"))  '***22/10/98
    End If
    
'Case "FN"
Case "F"
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH") - RS18("FN_PROF_CH"))  '***22/10/98
    End If
Case "NC"
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("NC_CH"))  '''''RRRR26799
''''' ADDED BY SANDHYA DATED 6.07.2012 FOR BP PATIENT RGJAY
Case "BP"
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH")
'If ncchg <> "I" Then '**22/10/98
'    DBGrid1.Columns(3) = 0
'Else                 '**22/10/98
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH") '**22/10/98
'End If '**22/10/98
Case Else
    MsgBox "NO MATCH ..."
End Select

If ncchg = "Z" Then ' added for z-series memos***4899

  DBGrid1.col = 3
Else
  DBGrid1.col = 4
End If   ' added for z-series memos***4899

DBGrid1.SetFocus


ElseIf ColIndex = 4 And ncchg <> "Z" Then 'ADDED EXTRA AND "NCCHG<> z" CONDITION 'ECTV
    CHKINDOCNAME
 
Else
Exit Sub
End If
Exit Sub
'DBGrid1.Row = DBGrid1.Row + 1
'DBGrid1.Col = 0
MSG:
If DBGrid1.col = 0 Then
   DBGrid1.col = 0
End If
DBGrid1.SetFocus
Exit Sub
Text22.Text = " PRESS CONTROL KEY TO COME OUT OF DBGRID .... and TO REFRESH TOTAL AMOUNT" '****9499
End Sub

Private Sub DBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, cancel As Integer)
ncchg = Left(DBGrid1.Columns(0), 1) ' added for z-series memos***4899

 If ColIndex = 3 Then
 If ncchg = "Z" Then
  Exit Sub
 Else
  MsgBox "YOU CANNOT MODIFY AMOUNT FOR " & CStr(DBGrid1.Columns(0)) & "   SERVICE"
  'KeyAscii = 0
  cancel = True
  Exit Sub
 End If
Exit Sub
End If  ' added for z-series memos***4899

'----------ECTV ADDED FROM HERE

 If ColIndex = 4 Or ColIndex = 5 Then
 If ncchg <> "Z" Then
  Exit Sub
 Else
  MsgBox "YOU CANNOT ENTER OR SELECT DOCTOR CODE AND NAME  FOR " & CStr(DBGrid1.Columns(0)) & "   SERVICE"
  'KeyAscii = 0
  cancel = True
  Exit Sub
 End If
Exit Sub
End If  ' added for z-series memos***4899









'----------ECTV ADDED UPTO HERE


End Sub




Private Sub DBGrid1_Click()
DBGrid1.Columns(1).Locked = True
'DBGrid1.Columns(3).Locked = True  ' commented for z-series memos***4899

End Sub
Private Sub DBGrid1_GotFocus()
DBGrid1.Columns(5).Locked = True
Text22 = " PRESS CONTROL KEY TO COME OUT OF DBGRID ....."
End Sub

Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo MSG
'AAAAAAAAAAAAAAAAAAAAAAA
' ADDED TO MAKE THE DOC CODE BLANK WHEN NO PROF CHARGES EXIST
If DBGrid1.col = 4 And KeyCode = 13 And Val(DBGrid1.Text) <> 0 Then
'    CHKINDOCNAME
    CHKDOCCODE
'Else
'    CHKINDOCNAME ' EVERY TIME CHECKING DOC CODE IS CORRECT OR NOT
End If
If DBGrid1.col = 5 And KeyCode = 13 Then
   DBGrid1.col = 4
   If Len(Trim(DBGrid1.Text)) = 0 Then
      CHKDOCCODE
   Else
      CHKINDOCNAME
   
   'DBGrid1.row = DBGrid1.row + 1
   'DBGrid1.Col = 0
    
'         Set RS18 = rdb.OpenResultset("select * from MA101 where ser_code = '" & DBGrid1.Columns(0) & "'", rdOpenKeyset, 3)
'opt = lbl_cat
'
'Select Case opt
'
'Case "B "
'   If RS18("B_PROF_CH") <> 0 Then
'      MsgBox "ENTER DOCTOR CODE....."
'      DBGrid1.Col = 4
'      DBGrid1.SelLength = 5
'      DBGrid1.SetFocus
'   Else
'       DBGrid1.row = DBGrid1.row + 1
'       DBGrid1.Col = 0
'   End If
'
'Case "A "
'     If RS18("A_PROF_CH") <> 0 Then
'      MsgBox "ENTER DOCTOR CODE....."
'      DBGrid1.Col = 4
'      DBGrid1.SelLength = 5
'      DBGrid1.SetFocus
'   Else
'       DBGrid1.row = DBGrid1.row + 1
'       DBGrid1.Col = 0
'   End If
'
'Case "D "
'    If RS18("D_PROF_CH") <> 0 Then
'      MsgBox "ENTER DOCTOR CODE....."
'      DBGrid1.Col = 4
'      DBGrid1.SelLength = 5
'      DBGrid1.SetFocus
'   Else
'       DBGrid1.row = DBGrid1.row + 1
'       DBGrid1.Col = 0
'   End If
'Case "F "
'    If RS18("FN_PROF_CH") <> 0 Then
'      MsgBox "ENTER DOCTOR CODE....."
'      DBGrid1.Col = 4
'      DBGrid1.SelLength = 5
'      DBGrid1.SetFocus
'   Else
'       DBGrid1.row = DBGrid1.row + 1
'       DBGrid1.Col = 0
'   End If
'
'Case Else
'    DBGrid1.row = DBGrid1.row + 1
'    DBGrid1.Col = 0
'End Select
   End If
End If

Exit Sub
MSG:
If DBGrid1.col = 0 Then
   DBGrid1.col = 0
End If
DBGrid1.SetFocus
Exit Sub

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 17 Then
txt_rem.Enabled = True
txt_rem.SetFocus
End If
End Sub

Private Sub exit1_Click()
'Call exit_prog
Unload Me
End
End Sub
Private Sub DBGrid1_LostFocus()
Dim tot As Long
Dim iRow As Integer
Text22 = " "
tot = 0
For iRow = 0 To mTotalRows&
  If Len(Trim(UserData(0, iRow))) > 0 Then '''''28041999
    If Not IsEmpty(UserData(0, iRow)) Then
    '' SANDHYA PROTOCOL PATIENT
     If PROTFLG = True Then
        TXTPRTTOT = CDbl(PTOT)
     End If
           tot = tot + CLng(UserData(3, iRow))
     End If
     '' END OF PROTOCOL
  End If
Next iRow

lbl_tot_amt = CDbl(Format(tot, "#,##,##0.00"))
'' SANDHYA PROTOCOL PATIENT
If PROTFLG = True Then
TXTNET = lbl_tot_amt - Val(TXTPRTTOT)
Else
TXTNET = lbl_tot_amt
End If
'' END OF PROTOCOL
End Sub
Private Sub delete_Click()
flag1 = 3
view_Click
btn_vis (0)
Enblbox (0) '19599
End Sub
Private Sub DBgrid1_KeyPress(KeyAscii As Integer)
On Error GoTo MSG
Dim rs_dcode As rdoResultset
'Dim ncchg As String ' commented for z-series memos***4899
NOCUR = 0
Dim opt As String
opttyp = lbl_typ
ncchg = Left(DBGrid1.Columns(0), 1) '***22/10/98
KeyAscii = Asc(UCase(Chr(KeyAscii)))


If KeyAscii = 13 Then
   If DBGrid1.col = 3 Or DBGrid1.col = 4 Or DBGrid1.col = 5 Then ' added for z-series memos***4899

    If DBGrid1.col = 4 Then
    CHKINDOCNAME
    'DBGrid1.row = DBGrid1.row + 1
    'DBGrid1.Col = 0
    'DBGrid1.SetFocus
    End If
Exit Sub
End If ' added for z-series memos***4899

If rs18flag = False Then
    RS18QR(0) = DBGrid1.Columns(0)
    Set RS18 = RS18QR.OpenResultset(1, 3)
Else
    RS18QR(0) = DBGrid1.Columns(0)
    RS18.MoreResults
    RS18.Requery
End If
'Set RS18 = rdb.OpenResultset("select * from MA101 where ser_code = '" & DBGrid1.Columns(0) & "'", rdOpenKeyset, 3) 'BILLOPT 61299 COMMENTED
If RS18.RowCount = 0 Then
    MsgBox " SERVICE CODE DOES NOT EXIST IN THE MASTER ..."
    DBGrid1.col = 0
    DBGrid1.SelStart = 0 ''21899
    DBGrid1.SelLength = 5
    DBGrid1.SetFocus
    Exit Sub '***23899
Else

DBGrid1.Columns(1) = RS18("ser_det")
EmptyGrid = 1
If dept = "BP" Then
    If SELCOLLECTION_FLAG = False Then
        SELCOLLECTION_QR(0) = Mid(DBGrid1.Columns(0), 6, 13)
        Set SELCOLLECTION_RS = SELCOLLECTION_QR.OpenResultset(1, 3)
        SELCOLLECTION_FLAG = True
    Else
        SELCOLLECTION_QR(0) = Mid(DBGrid1.Columns(0), 6, 13)
        SELCOLLECTION_RS.MoreResults
        SELCOLLECTION_RS.Requery
    End If
    If Not SELCOLLECTION_RS.EOF Then
        DBGrid1.Columns(2).Text = Trim(SELCOLLECTION_RS("COLLECTION"))
    End If
Else
    DBGrid1.Columns(2).Text = 1
End If
'opt = lbl_cat 'commented on 24599
''24599

'In cross-matrix table (bills/ma104) value
' for pat_cat="FN" and ch_code="FN".
If Trim(lbl_cat) = "F" Then
   lbl_cat = "F"
End If
If rs21flag = False Then
    RS21QR(0) = CASENO
    Set rs21 = RS21QR.OpenResultset(1, 3)
    rs21flag = True
Else
    RS21QR(0) = CASENO
    rs21.MoreResults
    rs21.Requery
End If
If rs21.RowCount <> 0 Then
    If rsdbtpflag = False Then
        RSDBTPQR(0) = Trim(rs21("bed_type"))
        RSDBTPQR(1) = Trim(lbl_cat)
        Set RSDBTP = RSDBTPQR.OpenResultset(1, 3)
        rsdbtpflag = True
    Else
        RSDBTPQR(0) = Trim(rs21("bed_type"))
        RSDBTPQR(1) = Trim(lbl_cat)
        RSDBTP.MoreResults
        RSDBTP.Requery
    End If
End If
'MsgBox RSDBTP.RowCount
RowCount = RSDBTP.RowCount
If RSDBTP.EOF = True Then
MsgBox "NO CORRESPONDING CHARGES AGAINST BED TYPE AND PATIENT CATEGORY", vbCritical
DBGrid1.col = 0
DBGrid1.SelLength = 4
DBGrid1.SetFocus
NOCUR = 1
Exit Sub
End If

'Set RSBDTP = rdb.OpenResultset("select ch_code from ma107 where pat_cat = '" & lbl_cat & "' and bed_type = '" & rs21("bed_type") & "'", 2, 3) 'BILLOPT 61299 COMMENTED
opt = RSDBTP("ch_code")
''24599
''' sandhya protocol patient starts
         
        PROTCASE_QR(0) = CASENO
        PROTCASE_QR(1) = MIRBNO
        
        Set PROTCASE_RS = PROTCASE_QR.OpenResultset(2, 3)
        PROTCASE_RS.Requery
        If PROTCASE_RS.RowCount > 0 Then
           PROTSERV_QR(0) = MIRBNO
           PROTSERV_QR(1) = Left(DBGrid1.Columns(0), 4)
           Set ZAS400DTTM = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1,CURRENT TIME TM1 FROM PATOOLSLIB.MA107") 'added for date
          ' ZAS400DTTM.Requery

           Set PROTSERV_RS = PROTSERV_QR.OpenResultset(2, 3)
           PROTSERV_RS.Requery
          If PROTSERV_RS.RowCount = 1 Then
            If Trim(PROTSERV_RS("FLD1")) = "E" Then
              DBGrid1.Columns(3) = PROTSERV_RS("ser_amt")
             
              PTOT = PTOT + DBGrid1.Columns(3)
              PROTFLG = True
            
            ElseIf Trim(PROTSERV_RS("FLD1")) = "I" Then
              PROTFLG = True
              Select Case opt

Case "C "
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH")
Case "B "
   If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("B_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("B_HOSP_CH") - RS18("B_PROF_CH"))  '***22/10/98
    End If
        
Case "A "
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("A_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("A_HOSP_CH") - RS18("A_PROF_CH"))  '***22/10/98
    End If
    
Case "D "
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("D_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("D_HOSP_CH") - RS18("D_PROF_CH"))  '***22/10/98
    End If
    
'Case "FN" '24599
Case "F " '24599
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH") - RS18("FN_PROF_CH"))  '***22/10/98
    End If
Case "NC"
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("NC_CH"))  '''''RRRR26799
''''' ADDED BY SANDHYA DATED 6.07.2012 FOR BP PATIENT RGJAY
Case "BP"
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH")
Case Else
    MsgBox "NO MATCH ..."
End Select
End If
Else
   NPROTFLG = True
            
Select Case opt

Case "C "
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH")
Case "B "
   If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("B_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("B_HOSP_CH") - RS18("B_PROF_CH"))  '***22/10/98
    End If
        
Case "A "
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("A_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("A_HOSP_CH") - RS18("A_PROF_CH"))  '***22/10/98
    End If
    
Case "D "
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("D_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("D_HOSP_CH") - RS18("D_PROF_CH"))  '***22/10/98
    End If
    
'Case "FN" '24599
Case "F " '24599
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH") - RS18("FN_PROF_CH"))  '***22/10/98
    End If
Case "NC"
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("NC_CH"))  '''''RRRR26799
'If ncchg <> "I" Then '**22/10/98
'    DBGrid1.Columns(3) = 0
'Else                 '**22/10/98
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH") '**22/10/98
'End If '**22/10/98

Case Else
    MsgBox "NO MATCH ..."
End Select
End If
Else
  NPROTFLG = True
''' END OF PROTOCOL
Select Case opt

Case "C "
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH")
Case "B "
   If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("B_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("B_HOSP_CH") - RS18("B_PROF_CH"))  '***22/10/98
    End If
        
Case "A "
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("A_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("A_HOSP_CH") - RS18("A_PROF_CH"))  '***22/10/98
    End If
    
Case "D "
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("D_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("D_HOSP_CH") - RS18("D_PROF_CH"))  '***22/10/98
    End If
    
'Case "FN" '24599
Case "F " '24599
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH") - RS18("FN_PROF_CH"))  '***22/10/98
    End If
Case "NC"
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("NC_CH"))  '''''RRRR26799
'If ncchg <> "I" Then '**22/10/98
'    DBGrid1.Columns(3) = 0
'Else                 '**22/10/98
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH") '**22/10/98
'End If '**22/10/98

Case Else
    MsgBox "NO MATCH ..."
End Select
End If
'If ncchg <> "Z" Then 'ADDED ONLY IF AND ENDIF 'ECTV

'DBGrid1.Columns(4) = lbl_dcode.Caption
'DBGrid1.Columns(5) = txt_unit_doc_name
'End If  'ADDED "AND ONLY IF AND ENDIF 'ECTV

DBGrid1.col = DBGrid1.col + 1
DBGrid1.SetFocus
'If ncchg <> "Z" Then 'ADDED ONLY IF AND ENDIF 'ECTV

DBGrid1.Columns(4) = lbl_dcode.Caption
DBGrid1.Columns(5) = txt_unit_doc_name
'End If  'ADDED "AND ONLY IF AND ENDIF 'ECT

'DBGrid1.Row = DBGrid1.Row + 1
 End If
 End If
Exit Sub
'MSG: MsgBox Error(Err) '" SERVICE / DOCTOR CODE DOES NOT EXIST IN THE MASTER ..."
MSG: MsgBox (" DOCTOR CODE DOES NOT EXIST ...")
'DBGrid1.SetFocus
End Sub
Private Sub flex2_Click()
On Error GoTo MSG
Dim opt As String
RSDNAMEQR(0) = flex2.Text 'BILLOPT 61299 ADDED
rs_dname.Requery 'BILLOPT 61299 ADDED
'Set rs5 = rdb.OpenResultset("select * from MA102 where doc_code = " & flex2.Text, 2, 3) 'or doc_code = " & Val(DBGrid1.Columns(4).Text), rdOpenKeyset, 3) ' 9/6/99 COMMENTED FROM 'OR DOCCODE
  'If rs5.RowCount <> 0 Then 'BILLOPT 61299 COMMENTED
    If rs_dname.RowCount <> 0 Then 'BILLOPT 61299 ADDED
      DBGrid1.Columns(5) = rs_dname("doc_name")
      DBGrid1.Columns(4) = rs_dname("doc_code")
 Else
   Exit Sub
 End If  '9/6/99 ADDED FROM EMPTY DOC CODE

flex2.Visible = False
DBGrid1.row = DBGrid1.row + 1
DBGrid1.col = 0
DBGrid1.SetFocus
a = 1
Exit Sub
MSG:  MsgBox Error(Err)

End Sub


Private Sub flex2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
flex2_Click
End If
End Sub

Private Sub Form_Unload(cancel As Integer)
'Dim QR As rdoQuery
'Dim RS As rdoResultset
'
'For Each RS In rdb.rdoResultsets
' RS.Close
'Next
'
'For Each QR In rdb.rdoQueries
' QR.Close
'Next

'RQ_SERIES.Close
'RS_SERIES.Close
'
'rq_updtSERIES.Close
'rq_reqdet.Close
'RS_reqdet.Close
'rq_updtreqdet.Close
'rq_updtdelreqdet.Close
'RQ_REQDETREQNO.Close
'RS_REQDETREQNO.Close
'requisitn_qr.Close
'requisitn_rs.Close
'UPDTA107_QR.Close
'InPt_ChkQr.Close
'RSINPT.Close
'
'REQUISITN_UPDT.Close
'RQ_reqdet1.Close
'Exist_ChkQr.Close
'rs22_CHK.Close
'
'RS_LVLQR.Close
'RS12QR.Close
'RS14QR.Close
'RS16QR.Close
'RS15QR.Close
'RS2QR.Close
'RS33QR.Close
'RS33QRCD.Close
'
'RS21QR.Close
'RS18QR.Close
'RSDBTPQR.Close
'RSDNAMEQR.Close
'RS5QR.Close
'Ma104_Qr.Close
'ma104_rs.Close
'RS_CR_QR.Close
'RS_TR_QR.Close
'RS44QR.Close
'addta107_QR.Close
'addta1071_QR.Close
'rs7qr.Close
'RS10QR.Close
'rs22qr.Close
'RS15QR1.Close
'RS151.Close
'
' rs_cr.Close
' rs_tr.Close
'rs30.Close
'rs31.Close
'rs1.Close
'rs33.Close
'rs22.Close
'rs3.Close
'rs4.Close
'rs5.Close
'rs6.Close
'rs7.Close
'rs8.Close
'rs10.Close
'rs12.Close
'RS11.Close
'rs13.Close
'RS15.Close
'RS14.Close
'RS16.Close
'RS17.Close
'rs2.Close
'RS20.Close
'RS18.Close
'RS19.Close
'RSDBTP.Close
'RS.Close
'rs21.Close

End Sub

Private Sub M_number_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0 'M_number1_LostFocus
End If
End Sub

Private Sub M_number_LostFocus()
''If flag = 1 Then
'' Exit Sub
''End If
'On Error GoTo msg:
''Dim st As String
''Dim I As Integer
''If flag = 2 Then
''ReDim UserData(0 To 5, 0 To 20)
''DBGrid1.Refresh
''End If
''DT = Mid(M_number, 2, 1)
'''M_number.Text = pad(6, "M", M_number, 1)
'' M_number.Text = padzero(6, "M", Mid(M_number, 3), 1)
''If M_number.Text <> " " Then
''
''
''        RS44QR(0) = M_number '"M-" & M_number
''        RS44QR(1) = "I"
''        Set rs10 = RS44QR.OpenResultset(2, 3)
''        If rs10.RowCount = 0 Then
''            MsgBox "NO CORRESPONDING MEMO EXISTS..."
''            Exit Sub
''        End If
''
'''        st = "select * from TA1071 where memo_no = '" & "M-" & M_number.Text & "'" 'BILLOPT 61299 COMMENTED
'''        Set RS7 = rdb.OpenResultset(st, rdOpenKeyset, 3) 'BILLOPT 61299 COMMENTED
''
'''         rs7qr(0) = "M-" & M_number.Text 'billopt 61299 added
'''         rs7.Requery 'billopt 61299 added
''
''
''        'com_mno.Locked = True
''       ' Set rs10 = rdb.OpenResultset("select * from TA107 where memo_no = '" & "M-" & M_number.Text & "'", rdOpenKeyset) 'billopt 61299 commented
''
''        'rs7qr(0) = "M-" & M_number.Text 'billopt 61299 added
''
''       ' rs7.Requery 'billopt 61299 added
''
''        me_date.Text = Format(rs10("memo_date"), "dd/mm/yyyy")
''        lbl_dno = rs10("dept_code")
''        lbl_cno = rs10("case_no")
''        'lbl_cat = rs10("MEMO_catg") '''CHECK FOR "memo_catg"'**commented 9/11/99 as there is no memo_catg fld in ta107
''        lbl_cat = rs10("MEMO_catg")
''        rs7qr(0) = M_number.Text '"M-" & M_number.Text 'billopt 61299 added
''        rs7.MoreResults
''        rs7.Requery 'billopt 61299 added
''        If rs7.RowCount = 0 Then
''        MsgBox "No Record found"
''        Exit Sub
''        End If
''        'txt_req_no = rs10("req_no")
''        txt_req_no.MaxLength = 13
''        txt_req_no.Text = rs7("req_no")
''        me_rdate.Mask = "##/##/####"
''            If IsDate(rs7("req_date")) Then
''                me_rdate.Text = Format(rs7("req_date"), "DD/MM/yyyy")
''            End If
''        lbl_dcode.Enabled = False
''        txt_unit_doc_name.Enabled = False
''        DBGrid1.Columns(0).Locked = True
''        DBGrid1.Columns(2).Locked = True
''        d_u_name.Visible = False
''        'Com_cno.Visible = False
''        txt_rem.Enabled = False
''        txt_doc_sec_name.Enabled = False
''        Call patient_det(lbl_cno)
''            If Not IsNull(rs10("doc_code")) Then
''                lbl_dcode.Caption = rs10("doc_code")
''                If rs2("pat_reg_ty") = "P" Then
''                    'Set rs1 = rdb.OpenResultset("select doc_name from MA102 where doc_code = " + lbl_dcode, rdOpenKeyset)  'billopt 61299 commented
''                    'txt_unit_doc_name = rs1("doc_name") 'billopt 61299 commented
''                     RSDNAMEQR(0) = lbl_dcode.Caption 'billopt 61299 added
''                     rs_dname.Requery 'billopt 61299 added
''
''
''                End If
''            Else
''                MsgBox (" Referred to unit ")
''            End If
''        txt_rem = rs10("rem")
''        'txt_c_cr = RS10("c_cr")
''        lbl_tot_amt = Format(rs10("memo_amt"), "#,##,##0.00")
''            If Not IsNull(rs10("d_s_name")) Then
''                    txt_doc_sec_name = rs10("d_s_name")
''            End If
''            If Not IsNull(lbl_dcode) Then
''                rs12.MoveFirst 'billopt 61299 addded
''                Do While rs12.EOF = False 'billopt 61299 addded
''                 If Trim(rs12("dept_code")) = Trim(lbl_dno) Then 'billopt 61299 addded
''                 lbl_dname = rs12("dept_name") 'billopt 61299 addded
''
''                  Exit Do
''                 Else
''                   rs12.MoveNext
''                 End If
''                 Loop 'billopt 61299 addded
''                'Set rs2 = rdb.OpenResultset("select dept_name from MA106 where dept_code = '" & lbl_dno & "'", rdOpenKeyset) 'billopt 61299 commented
''
''                'lbl_dname = rs2("dept_name")'billopt 61299 commented
''            End If
''
''
''
''
''
''Dim rrow
''If rs7.RowCount Then
'' rs7.MoveFirst
''For rrow = 0 To rs7.RowCount - 1
''     If Len(Trim(rs7(6).Value)) = 0 Then
''     UserData(0, rrow) = rs7(1).Value
''     Else
''     UserData(0, rrow) = rs7(1).Value & "/" & rs7(6).Value & "/" & rs7(7).Value & "/" & rs7(7).Value
''     End If
''        If Not IsNull(rs7(1).Value) Then
''            'Set rs13 = rdb.OpenResultset("select ser_det from MA101 where ser_code = '" & RS7(1).Value & "'", rdOpenKeyset, 3)' billopt commented
''             RS5QR(0) = rs7(1).Value 'billopt added
''             RS5QR(1) = "aaaaaaaaaa" 'billopt added
''             Set rs5 = RS5QR.OpenResultset(2, 3) 'billopt added
''             If rs5.RowCount > 0 Then
''             UserData(1, rrow) = rs5("ser_det") 'billopt added
''             End If
''            'UserData(1, rrow) = rs13("ser_det") 'billopt commeneted
''            UserData(2, rrow) = CStr(rs7(4).Value)
''            UserData(3, rrow) = CStr(rs7(2).Value)
''                If Not rs7(5).Value = 0 Then
''                    UserData(4, rrow) = CStr(rs7(5).Value)
''                    RSDNAMEQR(0) = CStr(rs7(5).Value) 'BILLOPT 61299 ADDED
''                    rs_dname.Requery 'BILLOPT 61299 ADDED
''                    'Set rs_dname = rdb.OpenResultset("select doc_name from ma102 where doc_code = " + CStr(rs7(4).Value), 2, 3) 'BILLOPT 61299 COMMENTED
''                    UserData(5, rrow) = rs_dname("DOC_NAME")
''                End If
''        End If
''   rs7.MoveNext
''Next rrow
''End If
''DBGrid1.Refresh
''End If
''new_view (lbl_cno)
'''If flag1 = 3 Then
'''Text22 = "press TAB to get the delete message...."
'''End If
''If flag1 = 3 Then
''    'Text22 = "press TAB to get the delete message...."
''    del
''    flag1 = 0
''    End If
''Exit Sub
''MSG:  MsgBox "ENTER VALID MEMO NUMBER TO VIEW "
''M_number.Enabled = False
'''com_mno.SetFocus
''
End Sub

Private Sub me_cno_GotFocus()
If Len(Trim(Com_dno.Text)) = 0 Then
Com_dno.SetFocus
End If

End Sub

Private Sub me_cno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  '  SendKeys "{TAB}"
'me_cno_LostFocus
KeyAscii = 0
'txt_req_no.SetFocus
'If RS15("c") = 1 Then
'txt_unit_doc_name.SetFocus
'End If
End If
End Sub
Private Sub me_cno_LostFocus()
On Error Resume Next
Dim c As Integer

'''' ADDED BY SANDHYA DATED 18.7.2012 FOR BP RGJAY PATIENT
BPPATFLAG = False

If flag = 1 Then
If Len(Trim(me_cno.Text)) <> 0 Then
me_cno = pad(5, "C", me_cno.Text, 3)
CASENO = Left(me_cno, 2) & "/" & Right(me_cno, 5)
 
'This query is used to get the latest admission details of the patient from TA105
'There had to be only one record where admission date is there and discharge intimation is
'not given
If rs15flag = False Then
    RS15QR(0) = CASENO
    Set RS15 = RS15QR.OpenResultset(1, 3)
    rs15flag = True
Else
 RS15QR(0) = CASENO
 RS15.MoreResults
 RS15.Requery
End If
 'if record is found---has to be usually 1-- then don't do anything
If RS15.RowCount = 0 Then
'This query is used to get The last admission date and last discharge date
'from table TA105
   If RS151flag = False Then
        RS15QR1(0) = CASENO
        Set RS151 = RS15QR1.OpenResultset(1, 3)
        RS151flag = True
   Else
        RS15QR1(0) = CASENO
        RS151.MoreResults
        RS151.Requery
   End If
   If RS151.RowCount <> 0 Then
      If IsNull(RS151(0)) = False Then
      MsgBox " THE PATIENT IS  ADMITTED ON        ' " & RS151(0) & "',     AND    GET DISCHARGE ON       '" & RS151(1) & "'                            GO TO DELAYED MEMO SCREEN  ENTER IT AGAIN ...... "
      CANCEL_Click
      New1.SetFocus
      Exit Sub
  Else
      MsgBox "THE PATIENT IS OUT PATIENT "
      me_cno.SelStart = 0 '***16/10/98
      me_cno.SelLength = 8 '***16/10/98
      me_cno.SetFocus '***16/10/98
      Exit Sub
    End If
End If
End If

'This query is used to find out the status of the bed if the patient is
'admitted and not given discharge intimation and not discharged for the
'supplied case number
If ma104_rsflag = False Then
    Ma104_Qr(0) = CASENO
    Set ma104_rs = Ma104_Qr.OpenResultset(1, 3)
    ma104_rsflag = True
Else
    Ma104_Qr(0) = CASENO
    ma104_rs.MoreResults
    ma104_rs.Requery
End If
'if the patient is admitted then the relaevant bed details
'should be available in MA104.Get the Bed status details.
If ma104_rs.RowCount <> 0 Then
   If Trim(ma104_rs("status")) = "V" Then
      MsgBox "BED IS SHOWN AS VACANT...WHEN PATIENT IS ALREADY ADMITTED ON THIS BED"
      btn_vis (0)
      Enblbox (0)
      New1.Enabled = True
      New1.SetFocus
      Exit Sub
   End If
'"DI" check is there because --delayed memos menu is to be used once discharge intimation is
'given to the patients
   If ma104_rs("status") = "DI" Then
      MsgBox "ALREADY DISCHARGE INTIMATION GIVEN SO NO MORE MEMOS CAN BE GENERATED..."
      btn_vis (0)
      Enblbox (0)
      New1.Enabled = True
      New1.SetFocus
      Exit Sub
   End If
End If
'-------- CHECKING FOR THE PATIENT IS INPATIENT AND NC CHANDU
            If RSINPTflag = False Then
                InPt_ChkQr(0) = CASENO
                Set RSINPT = InPt_ChkQr.OpenResultset(1, 3)
                RSINPTflag = True
            Else
                InPt_ChkQr(0) = CASENO
                RSINPT.MoreResults
                RSINPT.Requery
            End If
             ''''Set RSINPT = rdb.OpenResultset("select count(*)  from ta105 where case_no = '" & CASENO & "' AND DISC_DATE IS NULL AND ADMS_DATE IS NOT NULL", rdOpenKeyset, 3) '**ADDED ON 26/8/99
             If RSINPT.RowCount > 0 Then   'ADDED 23899
             ''' CHECKING FOR INPATIENTS OF CASH PAYING & NO CHARGE
                  RS2QR(0) = CASENO
                  rs2.MoreResults
                  rs2.Requery
              If rs2.RowCount > 0 Then
              
              '23/01/2002-parag
              'For All NC Patients-Inpatients/Outpatients-and irrespective of
              'whether they are CP/TA/CR-only OUT-PATIENT MEMOS are to be generated.
              If (rs2("PAT_TYPE") = "CP" Or rs2("PAT_TYPE") = "TA" Or rs2("PAT_TYPE") = "CR") And rs2("PAT_CATG") = "NC" Then
'               PATSTR = "CASE NO. " + CASENO + " IS  CASH PAYING  AND NC TYPE DO YOU WANT TO GENERATE MEMOS ?"
                MsgBox "CATEGORY OF THIS PATIENT IS ->NC AND IS AN IN-PATIENT" & " SO GENERATE ** OUT-PATIENT MEMO **", vbCritical, "         ----*** Exceptional Case ***---         "
                CANCEL_Click
                cleartxt
                'txt_unit_doc_name.Enabled = False
                Exit Sub
               End If
              End If
                
              End If




'--------
Call patient_det(CASENO)
txt_unit_doc_name.Enabled = t
If invcat = True Then
  Exit Sub
End If
'************
If flag = 1 Then
txt_req_no.ZOrder 1
''   RQ_REQDETREQNO(0) = CASENO
''   RQ_REQDETREQNO(1) = "IP"
''   RQ_REQDETREQNO(2) = lbl_dno
' remove the comments if u want all requisitions for all departments  chandu
   If RS_REQDETREQNOflag = False Then
        RQ_REQDETREQNO(0) = CASENO
        RQ_REQDETREQNO(1) = "IP"
        Set RS_REQDETREQNO = RQ_REQDETREQNO.OpenResultset(1, 1)
        RS_REQDETREQNOflag = True
   Else
        RQ_REQDETREQNO(0) = CASENO
        RQ_REQDETREQNO(1) = "IP"
        RS_REQDETREQNO.MoreResults
        RS_REQDETREQNO.Requery
    End If
If RS_REQDETREQNO.RowCount > 0 Then
    While Not RS_REQDETREQNO.EOF
     cbo_reqno.AddItem req_chglen(RS_REQDETREQNO("reqnno"))
     RS_REQDETREQNO.MoveNext
    Wend
End If
End If
'***************

If DOCFLAG = False And rs2("PAT_REG_TY") = "P" Then
    btn_vis (0)
    cleartxt
    New1.Enabled = True
    New1.SetFocus
    Exit Sub
End If

Call ud_name '9999999
    new_view (CASENO)
End If
Else
 cancel.SetFocus
End If
cbo_reqno.SetFocus
txt_unit_doc_name.Enabled = True
End Sub

Private Sub me_date_LostFocus()
If IsDate(me_date.Text) Then
    If CDate(me_date.Text) > Now Then
        MsgBox "Memo Date Cannot Be More Than Today's Date", vbExclamation, "Incorrect Data"
        me_date.SetFocus
        Exit Sub
    Else
        If CDate(me_date.Text) < "01/04/2001" Then
            MsgBox "Memo Date Cannot Be Less than 01/04/2001", vbExclamation, "Incorrect Data"
            me_date.SetFocus
            Exit Sub
        End If
    End If
Else
    MsgBox "Invalid Memo Date", vbCritical, "Invalid Date"
    me_date.SetFocus
    Exit Sub
End If


End Sub

Private Sub me_mno_GotFocus()
Text22 = "PRESS ENTER KEY ....."
End Sub
Private Sub me_mno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'Or KeyAscii = 9 Then
'me_mno_LostFocus
KeyAscii = 0
End If
End Sub


Private Sub me_mno_LostFocus()
'Dim c As Integer
'Dim CANCELFLG
'Text22 = " "
''There are two parts to this if statement-
''1> if memo number is given
''   1)memo number exits
''   2)memo number does not exist.
'
''2> memo number is not given.
''   That is the length of memo number is zero.
'If Len(Trim(me_mno.Text)) <> 0 Then
'    me_mno = pad(6, "M", me_mno, 1)
'    'This query is used to check if the supplied memo number
'    'already exists in the database.
'    'NOTE: it does not distinguish between out-patient/In-patient
'    'memo numbers. That is the same number cannot be used.
'
'    Exist_ChkQr(0) = "M-" & me_mno
'    Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
'    If rs22_CHK.RowCount <> 0 Then
'        CANCELFLG = MsgBox(" MEMO NUMBER ALREADY EXISTS ...... Do you want to cancel ?", vbYesNo) '****31399
'        If CANCELFLG = vbYes Then
'            'Cancel the current operation
'            CANCEL_Click
'            Exit Sub
'        ElseIf CANCELFLG = vbNo Then
'            'Go back to memo no and enter another number
'            me_mno.SetFocus
'            Exit Sub
'        End If
'    End If
'
'    'This query counts the number if records having the supplied memo number
'    rs22qr(0) = "M-" & me_mno
'    Set rs22 = rs22qr.OpenResultset(2, 3)
'
'    'if there are no records for the supplied memo number
'    If rs22("c") = 0 Then
'    'Com_dno.SetFocus
'    'goto memo date field
'       ' me_date.SetFocus
'    Else
'        MsgBox "MEMO NUMBER ALREADY EXISTS ...... "
'        me_mno.Mask = " "
'        me_mno.Text = " "
'        me_mno.Mask = "M-######"
'        me_mno.SelStart = 2 '***16/10/98
'        me_mno.SetFocus '***16/10/98
'        Exit Sub
'     End If
'Else
'    'if the length of memo number is zero-that is memo number is
'    'not given
'    CANCEL.Enabled = True
'    CANCEL.Visible = True
'    'cancel.SetFocus
'End If
End Sub
Private Sub me_rdate_LostFocus()
'Dim rdate  ' As Date'gmk comm 29-4-00
'Dim r1date  'As Date'gmk comm 29-4-00
Dim RDATE As String ' gmk add 29-4-00
Dim r1date As String 'gmk add 29-4-00

Dim D1 As Integer

Dim D2 As Integer
'Set RS15 = rdb.OpenResultset("select *  from ta105 where case_no = '" & CASENO & "' and disc_date is null", rdOpenKeyset, 3) '**COMM ON 26/8/99
If rs15flag = False Then
    RS15QR(0) = CASENO
    Set RS15 = RS15QR.OpenResultset(1, 3)
    rs15flag = True
Else
    RS15QR(0) = CASENO
    RS15.MoreResults
    RS15.Requery
End If
If me_rdate.Text <> "" Then
      RDATE = Left(me_rdate.Text, 2) & "/" & Mid(me_rdate.Text, 3, 2) & "/" & Right(me_rdate.Text, 4) 'Format(me_rdate.Text, "dd/mm/yyyy")
          If Not IsDate(RDATE) Then
           MsgBox "Not valid date!"
           me_rdate.SetFocus
           Exit Sub
           End If
           If RS15.RowCount <> 0 Then
       r1date = Format(RS15("adms_date"), "dd/mm/YYYY")
   '    MsgBox r1date
       'If rdate <= me_date And rdate >= r1date Then
       D1 = DateDiff("d", Format(RDATE, "dd/mm/yyyy"), Format(me_date, "dd/mm/yyyy"))
'       d2 = DateDiff("d", r1date, Format(rdate, "dd/mm/yyyy"))'**17/1198
       'd2 = DateDiff("d", RS15("ADMS_DATE"), Format(rdate, "dd/mm/yyyy")) '**17/1198 'gmk comm 29-4-00
       D2 = DateDiff("d", Format(RS15("ADMS_DATE"), "dd/mm/yyyy"), Format(RDATE, "dd/mm/yyyy")) '**17/1198'gmk add 29-4-00
             
       
       If D1 >= 0 And D2 >= 0 Then
          DBGrid1.SetFocus
          Exit Sub
       Else
       MsgBox "REQUISITION DATE SHOULD BE EITHER LESS THAN OR EQUAL TO CURRENT DATE AND GREATER THAN OR EQUAL TO ADMISSION DATE"
       me_rdate.SelStart = 0
       me_rdate.SelLength = 10
       me_rdate.SetFocus
       End If
       End If
End If
End Sub

Private Sub new1_Click()

Dim a, a1
On Error GoTo MSG

REQGOTFALG = False 'Added by Shraddha on 29/10/09 as told by Mrs.SRJoshi

EmptyGrid = 0 '***3499r
AmtNotEq = 0  ''R21899

flag1 = 0
flag = 1
Enblbox (1) '19599
commit.Enabled = True
cancel.Enabled = True
DBGrid1.Columns(0).Locked = False
DBGrid1.Columns(2).Locked = False
me_date.Enabled = True
txt_unit_doc_name.Enabled = False
d_u_name.Visible = True

cleartxt
me_mno.Visible = True
M_number.Visible = False  'ADDED ON 13/8/99
me_mno.ZOrder
' COMMENTED FOR AUTO GENERATION OF MEMO NUMBERS
'me_mno.SetFocus
'me_mno.Mask = "M-######"
' COMMENTED FOR AUTO GENERATION OF MEMO NUMBERS
me_cno.Mask = ">??/#####"
me_cno.Text = "__/_____"

me_rdate.Mask = "##/##/####"
d_u_name.Visible = True
me_cno.Visible = True
me_cno.Enabled = True
me_cno.ZOrder
lbl_cno.Visible = False
txt_doc_sec_name.Enabled = True
txt_rem.Enabled = True
ReDim UserData(0 To 5, 0 To 20)
DBGrid1.Refresh
Com_dno.Visible = True

me_date = Format(DATE11, "DD/MM/yyyy")
'COMMENTED FOR AUTO GENERATION OF MEMO NO
'me_mno.SelStart = 2
'me_mno.SetFocus
'COMMENTED FOR AUTO GENERATION OF MEMO NO.
rs12.MoveFirst
me2 = rs12("dept_name")
Com_dno.Text = me2
txt_doc_sec_name = usr_name
btn_vis (1)
Com_dno.SetFocus
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub CANCEL_Click()
On Error GoTo MSG:
'   If flag = 1 Then
'   rs4.AddNew
'   rs4.CancelUpdate
'   rs7.AddNew
'   rs7.CancelUpdate
'   Else
'    rs4.Edit
'    rs4.CancelUpdate
'    rs7.Edit
'    rs7.CancelUpdate
'    End If
    cleartxt
        
     btn_vis (0)
     Enblbox (0) '19599
     Back.SetFocus
    'btn_vis (0)
    'Combo1.Clear
    Exit Sub
MSG:      MsgBox Error(Err)
End Sub

    Sub cleartxt()
    lbl_typ = " "
'    Com_dno = "  "
    me_mno.Mask = "  "
    lbl_age = " "
    lbl_cat = " "
    lbl_sex = " "
    lbl_wf_bed = " "
    LBL_NAME = " "
    lbl_dname = " "
    lbl_cr_card_no = " "
lbl_tr_no = " "
    'txt_c_cr = " "
    'd_u_name = " "
    lbl_tot_amt = " "
    lbl_dcode = " "
    txt_req_no = ""
    'd_u_name = "  "
    'me_rdate.Mask = " "
    me_cno.Mask = ""
    me_cno.Text = ""
    me_rdate.Mask = ""
    me_rdate.Text = ""
    ''ADDED BY SANDHYA PROTOCOL PATIENT
    PTOT = 0
    TXTPRTTOT = " "
    TXTNET = " "
    ''' END OF PROTOCOL
    lbl_dno = " "
    Text22 = " "
    txt_doc_sec_name = " "
    txt_unit_doc_name = " "
    txt_rem = "  "
    DBGrid1.ClearFields
    txt_req_no.Text = ""
    cbo_reqno.clear
    'DBGrid1.BackColor = RGB(0, 0, 0)
    End Sub

Sub commit_Click()
Dim i
Dim zconfirm As Integer
Dim zvalidid As Variant
Dim ZMAX As Double
Dim ztotser As Double
Dim SR_AMT As Double
ztotser = 0
FLGTA = 0
FLGCR = 0
'''' smart card patient flag added by sandhya
FLGSM = 0
'********************************20499
' COMMENTED FOR AUTO GENERATION OF MEMO NO.
'''If Len(Trim(me_mno.Text)) = 0 Then ''19599
'''    MsgBox "Memo number can't be empty !", vbCritical
'''    me_mno.SetFocus
'''    Exit Sub
'''End If ''19599
' COMMENTED FOR AUTO GENERATION OF MEMO NO.

'If ((MsgBox("Do you want to save") + vbYesNo) = vbNo) Then exitsub
'' BIJESH DUPLICATE MEMOS
Dim iRow As Integer
memoduplicate = False
'' END
If EmptyGrid = 0 Then
MsgBox "ATTEMPT TO SAVE EMPTY MEMO"
Exit Sub
End If
'********************************20499


'26/10/2005--BIJESH
'PatchUp for checking if the same Service code has been generated for the patient
'in the same day.
'''' added by Shraddha dated 30.8.2010 \
'''' double memo check for CR/TA patient
If Trim(lbl_typ) = "CR" Or Trim(lbl_typ) = "TA" Then

For iRow = 0 To mTotalRows&
  If Len(Trim(UserData(0, iRow))) > 0 Then '''''28041999
    If Not IsEmpty(UserData(0, iRow)) Then
       zgetSerCodes(0) = CASENO
       zgetSerCodes(1) = ZAS400DTTM(0)
       zgetSerCodes(2) = Trim(UserData(0, iRow))
       zrsgetSercodes.Requery
       
       '***** Shraddha 16/08/10 *****
       qrser_ref(0) = Trim(UserData(0, iRow))
       Set rsser_ref = qrser_ref.OpenResultset(1, 3)
       
       
       If zrsgetSercodes(0) > 0 Then
       
       If rsser_ref.RowCount = 1 Then
            memoduplicate = False
        Else
               ''' ADDED BY SANDHYA DATED 29.7.2010
         ''' TO STOP DOUBLE MEMOS
         ' i = MsgBox("The Service code  " & Trim(UserData(0, iRow)) & " has already been generated " & zrsgetSercodes(0) & " times with total " & zrsgetSercodes(1) & " quantities for this Patient Today..." & Chr(13) & "Do You Wish to Proceed.... ????", vbYesNo + vbDefaultButton2, "Confirmation Required..??")
          'MsgBox ("The Service code  " & Trim(UserData(0, iRow)) & " has already been generated " & zrsgetSercodes(0) & " times for this Patient Today...")
          
          ''' ADDED BY SANDHYA DATED 29.7.2010
       ''' TO STOP DOUBLE MEMOS
         ' i = MsgBox("The Service code  " & Trim(UserData(0, iRow)) & " has already been generated " & zrsgetSercodes(0) & " times with total " & zrsgetSercodes(1) & " quantities for this Patient Today..." & Chr(13) & "Do You Wish to Proceed.... ????", vbYesNo + vbDefaultButton2, "Confirmation Required..??")
          MsgBox ("Service code " & Trim(UserData(0, iRow)) & " already been generated " & zrsgetSercodes(0) & " time for this Patient Today, If you wish to generate please contact Opd Supervisor/MSOffice/Head of Dept")
          
          If Trim(USRGRP) <> "JAO/OPD_SUPVSR" Then
             MsgBox ("Memo cannot be generated twice")
             memoduplicate = True
             FLGCR = 0
             btn_vis (0)
             Exit Sub
          Else
            'Dim mpass As String
            Me.Hide
            Frmpass.ZOrder
            DoEvents
            Frmpass.show MODAL
          End If
       End If
'          I = MsgBox("The Service code  " & Trim(UserData(0, iRow)) & " has already been generated " & zrsgetSercodes(0) & " times with total " & zrsgetSercodes(1) & " quantities for this Patient Today..." & Chr(13) & "Do You Wish to Proceed.... ????", vbYesNo + vbDefaultButton2, "Confirmation Required..??")
'          If I = vbNo Then
'             memoduplicate = True
'          End If
       End If
    End If
  End If
Next iRow
End If ' (CHECK OF DOUBLE MEMO END)
If memoduplicate = True Then
   Exit Sub

Else

'********************************20499
'FOR MEMO SUN BLANK
If Len(Trim(lbl_tot_amt.Caption)) <> 0 Then     '1
'If rs2("PAT_TYPE") = "CR" Then  COMMENTED BY SEKHAR FOR DIS PURPOSE
If SEL_M3ICU_FLG = False Then   '5
            Set SEL_M3ICU_RS = SEL_M3ICU_QR.OpenResultset(1, 3)
            SEL_M3ICU_FLG = True
        Else
            SEL_M3ICU_RS.MoreResults
            SEL_M3ICU_RS.Requery
        End If                          '5
         
        If SEL_M3ICU_RS.RowCount = 0 Then
            GoTo BYPASS
        End If
         
   If Trim(SEL_M3ICU_RS("ipadd")) <> Trim(ipadd) Then
BYPASS:
If Trim(lbl_typ) = "CR" And BPPATFLAG = False Then   '2
   If rs_crflag = False Then   '3
        RS_CR_QR(0) = CASENO
        Set rs_cr = RS_CR_QR.OpenResultset(1, 3)
        rs_crflag = True
   Else
        RS_CR_QR(0) = CASENO
        rs_cr.MoreResults
        rs_cr.Requery
  End If                        '3
 '25/10/2001-PARAG
 'Either the credit facility will be upto a certian amount
 'or upto a certain date. So it checks if the MAX_AMT is 0
 '(either because it may be upto a certain DATE or the amount has been exhaused and
 'became zero(0).
        If rs_cr.RowCount > 0 Then     '4
        
         ZMAX = CDbl(rs_cr("MAX_AMT"))
         'CHECK IF THE CREDIT LIMIT IS UPTO A MAX VALUE / CREDIT CARD DATE
         If ZMAX = 0 Then      '6
            If DATE11 > Format(rs_cr("DATE_MAX"), "DD/MM/YYYY") Then
                MsgBox "Credit Card Has Expired On : " & rs_cr("date_max"), vbCritical, "Credit Card Status"
                CANCEL_Click
                Exit Sub
            End If
        End If                 '6
        If CDbl(rs_cr("max_amt")) <> 0 Then    '7
                    If CDbl(Len(Trim(rs_cr("SER_AMT")))) = 0 Then
                        SR_AMT = 0
                    Else
                        SR_AMT = CDbl(rs_cr("SER_AMT"))
                    End If
         Else
         '25/10/2001-PARAG
'            rs_cr ("date_max") > Format(Date, "dd/mm/yyyy")
'                MsgBox "Credit Card Date Exceeds Maximum Limit", vbCritical, "Credit Card Status"
'                Exit Sub
         End If                                 '7
'FFFFFFFFFFFFFFFFFFFF
'                 If (SR_AMT + lbl_tot_amt) > (0.9 * rs_cr("MAX_AMT")) Then
'                    MsgBox "SERVICE AMOUNT EXCEEDS 90% OF MAXIMUM AMOUNT ; MEMO CANNOT BE ISSUED "
'                    MsgBox "COMPANY BALANCE AMOUNT = " & (rs_cr("MAX_AMT") - SR_AMT)
'                    FLGCR = 0
'                    btn_vis (0)
'                    Enblbox (0) '19599
'                    Exit Sub
'                 Else
'                    FLGCR = 1
'                 End If
        If rs_cr("max_amt") <> 0 Then   '8
                 If (CDbl(SR_AMT) + CDbl(lbl_tot_amt)) > (0.6 * CDbl(rs_cr("MAX_AMT"))) Then ' added on ms request '9
                     'If Trim(SEL_M3ICU_RS("IPADD")) <> ipadd Then
                      per = (CDbl(SR_AMT) + CDbl(lbl_tot_amt)) * 100 / CDbl(rs_cr("MAX_AMT"))
                      MsgBox "SERVICE AMOUNT EXCEEDS " & Chr(13) & "                          " & per & "%" & Chr(13) & "      OF MAXIMUM AMOUNT =" & rs_cr("MAX_AMT") & Chr(13) & "COMPANY BALANCE AMOUNT = " & (rs_cr("MAX_AMT") - SR_AMT)
                      FLGCR = 1
                     'nd If
                      If (CDbl(SR_AMT) + CDbl(lbl_tot_amt)) > (1 * CDbl(rs_cr("MAX_AMT"))) Then   '11
                      '  If Trim(SEL_M3ICU_RS("IPADD")) <> ipadd Then
                          i = MsgBox("SERVICE AMOUNT EXCEEDS   " & per & "%" & Chr(13) & " EXCEDING MEMO AMOUNT  = " & (lbl_tot_amt - (rs_cr("MAX_AMT") - SR_AMT)) & Chr(13) & "GET PERMISSION OF HIGHER OFFICIALS TO ISSUE MEMO", vbYesNo + vbDefaultButton2, "COMMIT BOX") ' ADDED FOR THE 100% OF MAX_AMT
                          If i = vbYes Then
                                FLGCR = 1
                          Else
                                FLGCR = 0
                                btn_vis (0)
                                Exit Sub
                          End If
                       ' End If
                      End If '11
                 Else
                    FLGCR = 1
                 End If      '9
          End If             '8
 'FFFFFFFFFFFFFFFFFFF
        Else
           Set RS = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 FROM PATOOLSLIB.MA107") 'added for date
            If RS.RowCount > 0 Then
            If rs_cr("date_max") <= RS("dt1") Then
              'If Format(rs_cr("date_max"), "dd/mm/yyyy") < Format(me_date, "dd/mm/yyyy") Then '***21-4-00 ADDED FOR DATE
                 'If Format(rs_cr("date_max"), "YYYY-MM-DD") < GET_REVDATE(me_date) Then '***21-4-00 ADDED FOR DATE
                 'If rs_cr("date_max") < CDate(me_date) Then'***21-4-00 COMM FOR DATE
                       MsgBox "VALID DATE EXCEEDED ...."
                       FLGCR = 0
                       btn_vis (0)
                       Enblbox (0) '19599
                       Exit Sub
                End If
                MsgBox "MEMO CAN BE SAVED"
              End If
        End If   '4
'ElseIf rs2("PAT_TYPE") = "TA" Then
ElseIf Trim(lbl_typ) = "TA" Then

' Padmakar - (TMH-ACTREC) joint company account details
  Call TMH_ACTREC(CASENO)

'(*********************************)START
If rs_trflag = False Then    '12
    RS_TR_QR(0) = CASENO
    Set rs_tr = RS_TR_QR.OpenResultset(1, 3)
    rs_trflag = True
Else
    RS_TR_QR(0) = CASENO
    rs_tr.MoreResults
    rs_tr.Requery
End If                      '12
                    If Len(Trim(rs_tr("SER_AMT"))) = 0 Then
                        SR_AMT = 0
                    Else
                        SR_AMT = CDbl(rs_tr("ser_amt"))
                    End If
                    lbl_tot_amt.Refresh
                        iRow = 0
             If UserData(0, iRow) <> "Z100" Then       '13
                If (CDbl(SR_AMT) + CDbl(lbl_tot_amt)) > (0.9 * CDbl(rs_tr("NET_AMT"))) Then    '14
                    MsgBox "MEMO AMOUNT EXCEEDS 90% OF MAXIMUM AMOUNT ; MEMO CANNOT BE ISSUED "
                    lbl_tot_amt.Refresh
                    ztotser = CDbl(rs_tr("ser_amt")) + CDbl(lbl_tot_amt.Caption)
                    MsgBox "TRUST BALANCE AMOUNT = " & (CDbl(rs_tr("NET_AMT")) - CDbl(ztotser))
                   '***************************
                   '20/02/2002-Parag
                   'Even if the amount exceeds 90% of the total deposited
                   'amount but the balance after reducing the new memo amt
                   'is more than Rs100 the memo can be saved.
                   lbl_tot_amt.Refresh
                   If (rs_tr("NET_AMT") - (ztotser)) >= 100 Then    '15
                       FLGTA = 1
                       MsgBox "MEMO IS TO BE SAVED"
                   Else
                    '----------------------------------
                    'Even if 90% amount is exceeded for Group 3 (Accounts Group)
                    'for users listed in matchup function can save the memo
                    'on the basis of a sanction given by Authorised personnel.
                    '06/08/2001-Parag
                    
                     zvalidid = MatchUp(txt_doc_sec_name.Text)
                     If IsNull(zvalidid) Then          '16
                       FLGTA = 0
                       btn_vis (0)
                       Exit Sub
                     Else
                      zconfirm = MsgBox("GET PERMISSION OF HIGHER OFFICIALS TO ISSUE MEMO", vbYesNo + vbDefaultButton2, "COMMIT BOX")
                       If zconfirm = vbYes Then         '17
                          FLGTA = 1
                       Else
                          FLGTA = 0
                          btn_vis (0)
                          Exit Sub
                       End If                           '17
                    End If                              '16
                  End If 'for Rs100  condition          '15
                '----------------------------------
                Else
                FLGTA = 1
                MsgBox "MEMO IS TO BE SAVED"
                End If                                  '14
              Else
                If (ztotser) > (rs_tr("NET_AMT")) Then
                    MsgBox " Service amount exceeds balance amount", vbInformation, "Amount Exceeds"
                    FLGTA = 0
                    btn_vis (0)
                    Exit Sub
                Else
                    FLGTA = 1
                    MsgBox "MEMO IS TO BE SAVED"
                End If
              
            End If                         '13


'(******************************) END
End If                          '2
End If    'M3
End If

'''' ADDED BY SANDHYA FOR SMART CARD PATIENT
    If rs2("PAT_TYPE") = "CP" And scard = True Then
       Dim RQSCARDACC As rdoResultset
       Dim RQSMOSTCASE As rdoResultset
       Set RQSCARDACC = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.SCARDACC WHERE CASE_NO = '" & CASENO & "'", 2, 1)
       If RQSCARDACC.RowCount = 1 Then
          If Len(Trim(RQSCARDACC("SER_AMT"))) = 0 Then
             SR_AMT = 0
          Else
             SR_AMT = RQSCARDACC("SER_amt")
             NET_AMT = RQSCARDACC("DEPO_AMT")
          End If
          
          If NET_AMT - (SR_AMT + lbl_tot_amt) < 0 Then
            ' Set RQSMOSTCASE = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.SMOSTCASES WHERE CASE_NO ='" & CASENO & "' AND ALLOWYN = 'Y' AND PERM_AMT > 0 ", 1, 1)
            ' If RQSMOSTCASE.RowCount = 1 Then
                MsgBox "OVERDUE PAYMENT, TOP UP YOUR SMART CARD ACCOUNT, BALANCE IS BELOW ZERO", vbCritical
                
          ElseIf NET_AMT - (SR_AMT + lbl_tot_amt) < 5000 Then
             MsgBox "INSUFFICIENT BALANCE, PLEASE TOP UP YOUR SMART CARD ACCOUNT ", vbCritical
          Else
             MsgBox "SMART CARD BALANCE AMOUNT = " & (RQSCARDACC("DEPO_AMT") - (SR_AMT))
           
          End If
           FLGSM = 1
           MsgBox "MEMO IS TO BE SAVED"
       End If
    End If
    '''' END OF SMART CARD


'1
'' sandhya protocol patient starts
If PROTFLG = True Then
   PROTCASE_QR(0) = CASENO
   PROTCASE_QR(1) = MIRBNO
   Set PROTCASE_RS = PROTCASE_QR.OpenResultset(2, 3)
   If PROTCASE_RS.RowCount > 0 Then
      If Len(Trim(PROTCASE_RS("SER_AMT"))) = 0 Then
         SR_AMT = 0
      Else
         SR_AMT = CDbl(PROTCASE_RS("ser_amt"))
      End If
      TXTPRTTOT.Refresh
      If (SR_AMT + CDbl(prottot)) > (0.75 * PROTCASE_RS("TOTALAMT")) Then
                    MsgBox "MEMO AMOUNT EXCEEDS 75% OF MAXIMUM AMOUNT ; MEMO CANNOT BE ISSUED "
                    ztotser = PROTCASE_RS("ser_amt") + CDbl(TXTPRTTOT)
                    MsgBox "PROTOCOL BALANCE AMOUNT = " & (PROTCASE_RS("TOTALAMT") - (ztotser))
   
         prottot.Refresh
         If (PROTCASE_RS("TOTALAMT") - (ztotser)) >= 100 Then
             FLGTA = 1
             MsgBox "MEMO IS TO BE SAVED"
         Else
             zvalidid = MatchUp(txt_doc_sec_name.Text)
             If IsNull(zvalidid) Then
                 FLGTA = 0
                 btn_vis (0)
                 Exit Sub
              Else
                 zconfirm = MsgBox("GET PERMISSION OF HIGHER OFFICIALS TO ISSUE MEMO", vbYesNo + vbDefaultButton2, "COMMIT BOX")
                 If zconfirm = vbYes Then
                    FLGTA = 1
                  Else
                    FLGTA = 0
                    btn_vis (0)
                    Exit Sub
                  End If
              End If
         End If 'for Rs100  condition
                '----------------------------------
      Else
         FLGTA = 1
         MsgBox "MEMO IS TO BE SAVED"
      End If
   Else
              
   End If
 End If
'' protocol patient ends

On Error GoTo MSG:
i = MsgBox(" Do you want to save ? ", vbYesNo, "COMMIT BOX")
''added by mono
 'If Trim(cbo_reqno.Text) <> "" Then
 If cbo_reqno.ListCount <> 0 Then
    MsgBox " Please select all the Procedure from the Combo Box !!", vbInformation, "MEMO"
    cbo_reqno.SetFocus
    Exit Sub
 End If
If i = vbYes Then
   add_edit
End If
Me.Visible = True
btn_vis (0)
Enblbox (0) '19599
print1.SetFocus
       Exit Sub
MSG:
       MsgBox Error(Err)
End If  '' (BIJESH DUPLICATE MEMOS END)
End Sub


Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
'Dim iCol As Integer
'
'mTotalRows = mTotalRows + 1
'ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
'NewRowBookmark = mTotalRows - 1
'
'For iCol = 0 To UBound(UserData, 1)
'    If Not IsNull(RowBuf.Value(0, iCol)) Then
'        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
'    Else
'
'       UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
'    End If
'Next iCol

End Sub

Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub


' This subroutine is called, any time the DBGrid wants to display
' new data.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid is requesting rows so give them to it

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' If StartLocation is Null then start reading at the end
' or beginning of the data set.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Find the position to start reading based on the
    ' StartLocation bookmark and the iIncr variable
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfer data from our data set array to the RowBuf object
' which DBGrid uses to display the data
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Set bookmark using CurRow& which is also our
    ' array index
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub
Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer

For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub
Private Sub Form_Load()
NOCUR = 0
Dim i As Integer
flag1 = 0

'26/10/2005--BIJESH
Dim str As String
memoduplicate = False
'This query is used to get the AS400 Date and Time
Set ZAS400DTTM = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1,CURRENT TIME TM1 FROM PATOOLSLIB.MA107") 'added for date

Set RS = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 FROM PATOOLSLIB.MA107") 'added for date
DATE11 = RS("DT1")

'BIJESH Querry to get the details of Service Codes against Memos generated for the Patient in the same day.
str = "SELECT COUNT(B.SER_CODE),SUM(B.QTY) FROM PATOOLSLIB.TA107 A,PATOOLSLIB.TA1071 B WHERE A.MEMO_NO=B.MEMO_NO AND A.CASE_NO=? AND A.MEMO_DATE=? AND B.SER_CODE=?"
Set zgetSerCodes = rdb.CreateQuery("zgetSerCodes12S", str)
zgetSerCodes(0) = "TT/00300"
zgetSerCodes(1) = "2000-01-01"
zgetSerCodes(2) = "A001"
Set zrsgetSercodes = zgetSerCodes.OpenResultset(2, 1)

'' END OF MEMOS

'Main
If ipmemoflag = False Then
Set RS = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 FROM PATOOLSLIB.MA107") 'added for date
DATE11 = RS("DT1")


''29/10/03--MONO
 ''******FOR BP CROSSMATCH**********
 str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=? AND PRO_CODE='H002'"
 Set SELH002_QR = rdb.CreateQuery("SELH002_STR11", str)
 
str = "SELECT COLLECTION FROM DISLIB.DMULPROCE WHERE REQNO=? AND PRO_CODE='H002'"
Set SELCOLLECTION_QR = rdb.CreateQuery("SELCOLLECTION_STR", str)
 
 ''******FOR BP CROSSMATCH**********
''*********************M3ICU***************
    str = "SELECT * FROM DISLIB.DISICU WHERE M1='M3'"
    Set SEL_M3ICU_QR = rdb.CreateQuery("SEL_M3ICU_QR", str)
''*********************M31CU***************

'Set RS_LVLQR = rdb.CreateQuery("IPRS_LVLQ", "select * from TOOLSADMN.scrn_auth1 where TRIM(USR_GRP) = ? AND TRIM(SCRN_ID) = ?")  'BILLOPT 6/12/99

'LVL_BTN_ENABL '**1/7/99
'NEW_LVL_BTN
'delete.Visible = False
AmtNotEq = 0
'ReDim UserData(0 To 5, 0 To 20)
'mTotalRows& = 19
ReDim UserData(0 To 5, 0 To 12)
mTotalRows& = 12

Dim oldcnt As Integer, newcnt As Integer
'com_mno.Visible = False
'lbl_dno.Visible = False
lbl_cno.Visible = False
flag = 0
flex1.Visible = False
flex2.Visible = False
a = 1
'*** COMMENTED BY SEKHAR
Set RS12QR = rdb.CreateQuery("IPRS12Q", "SELECT DEPT_CODE,DEPT_NAME FROM PATOOLSLIB.MA106  ") 'BILLOPT 61299 ADDED

'Set rs12 = rdb.OpenResultset("select * from MA106", rdOpenKeyset, 1) ''19599\ 'BILLOPT 61299 COMMENTED
If rs12flag = False Then
    Set rs12 = RS12QR.OpenResultset(1, 3)
    rs12flag = True
Else
    rs12.MoreResults
    rs12.Requery
End If
If rs12.RowCount > 0 Then
While Not rs12.EOF ''19599 MOVED FROM NEW1_CLICK
 'Com_dno.AddItem rs12("dept_code")
 Com_dno.AddItem rs12("dept_name")
 rs12.MoveNext
 Wend
End If
'*** COMMENTED BY SEKHAR
'Set rs7 = rdb.OpenResultset("select * from TA1071", 2, 3)'**COMMENTED ON 9/9/99

'Me.Show
oldcnt = DBGrid1.Columns.count
newcnt = 0

' Remove old columns
For i = DBGrid1.Columns.count - 1 To 0 Step -1
    DBGrid1.Columns.Remove i
Next i

' Add new columns
For i = 0 To 5
    DBGrid1.Columns.add newcnt
    
    newcnt = newcnt + 1
    
Next i
DBGrid1.Columns(0).Width = 3700
    DBGrid1.Columns(0).Caption = "CODE/REQ.NO/REQ.DATE/REQ.CATG"
    DBGrid1.Columns(0).Visible = True
    DBGrid1.Columns(1).Width = 4000
    DBGrid1.Columns(1).Caption = "SERVICES/PROCEDURES/INVESTIGATIONS"
    DBGrid1.Columns(1).Visible = True
    DBGrid1.Columns(2).Width = 900
    DBGrid1.Columns(2).Caption = "QUANTITY"
    DBGrid1.Columns(2).Visible = True
    DBGrid1.Columns(3).Caption = "AMOUNT"
    DBGrid1.Columns(3).Visible = True
    DBGrid1.Columns(4).Width = 1500
    DBGrid1.Columns(4).Caption = "DOCTOR CODE"
    DBGrid1.Columns(4).Visible = True
    DBGrid1.Columns(5).Width = 3000
    DBGrid1.Columns(5).Caption = "DOCTOR NAME"
    DBGrid1.Columns(5).Visible = True
    DBGrid1.Columns(0).Button = True
 
 Set RS14QR = rdb.CreateQuery("IPRS14Q", "select U_D,UD_NAME from PATOOLSLIB.MA1031 WHERE U_D=?") 'BILLOPT 61299 ADDED
 'RS14QR(0) = "A"
 'Set RS14 = rdb.OpenResultset("select ud_name from MA1031", rdOpenKeyset) 'BILLOPT 61299 COMMENTED
' Set RS14 = RS14QR.OpenResultset(2, 3) 'BILLOPT 61299 ADDED
 
 'Set RS16QR = rdb.CreateQuery("RS16Q", "select DOC_name from MA102") 'BILLOPT 61299 COMMENTED
 'Set RS16 = RS16QR.OpenResultset(2, 3) 'BILLOPT 61299 COMMENTED
 'Set RS16 = rdb.OpenResultset("select doc_name from MA102", rdOpenKeyset) 'BILLOPT 61299 COMMENTED
 
 '****************************************************************************************************************************
 'This query is used to get the latest admission details of the patient from TA105
 Set RS15QR = rdb.CreateQuery("IPRS15Q", "select CASE_NO,ADMS_DATE   from PATOOLSLIB.ta105 where case_no = ? and disc_date is null AND ADMS_DATE IS NOT NULL") 'BIILLOPT 61299 ADDED FOR INPATIENT CHECK
' RS15QR(0) = ">??/#####"
' Set RS15 = RS15QR.OpenResultset(2, 3)
 '****************************************************************************************************************************
 
 '****************************************************************************************************************************
 'This query is used to get The last admission date and last discharge date
 'from table TA105
 Set RS15QR1 = rdb.CreateQuery("IPRS15Q1", "select MAX(ADMS_DATE),MAX(DISC_DATE)   from PATOOLSLIB.ta105 where case_no = ? ") 'BIILLOPT 61299 ADDED FOR INPATIENT CHECK
' RS15QR1(0) = ">??/#####"
' Set RS151 = RS15QR1.OpenResultset(2, 3)
 '****************************************************************************************************************************
 
 '****************************************************************************************************************************
 'This query is used to get the details of patients from Ta101
 Set RS2QR = rdb.CreateQuery("IPRS2Q", "select * from PATOOLSLIB.TA101 where case_no = ?") 'BILLOPT 61299 ADDED FOR PATIENT DET
' RS2QR(0) = "BN/00000"
'    Set rs2 = RS2QR.OpenResultset(2, 3)
 '****************************************************************************************************************************
 
 '****************************************************************************************************************************
 'This query gets the doctor code from DOCVIEW
 Set RS33QR = rdb.CreateQuery("IPRS33", "SELECT DOC_CODE FROM PATOOLSLIB.DOCVIEW WHERE UD_NAME = ?") 'BILLOPT 61299 ADDED FETCHING FROM VIEW
 
 Set RS33QRCD = rdb.CreateQuery("IPRS33QCD", "SELECT UD_NAME FROM PATOOLSLIB.DOCVIEW WHERE DOC_CODE = ?") 'BILLOPT 61299 ADDED FETCHING FROM VIEW
 
 '****************************************************************************************************************************
 
 '****************************************************************************************************************************
 Set RS21QR = rdb.CreateQuery("IPRS21Q", "select ward_name,bed_type,bed_no,FLOOR from PATOOLSLIB.MA104 where case_no = ? ") 'BILLOPT 61299 ADDDED FOR BED DETAILS IN VIEW
 '****************************************************************************************************************************
 
 '****************************************************************************************************************************
 Set RS18QR = rdb.CreateQuery("IPRS18Q", "select * from PATOOLSLIB.MA101 where ser_code = ?") 'BILLOPT 61299 ADDED
' RS18QR(0) = "B001" 'BILLOPT 61299 ADDED
' Set RS18 = RS18QR.OpenResultset(2, 3) 'BILLOPT 61299 ADDED
 '****************************************************************************************************************************

'Get the details from mapping table : SHI1.SERVICED FOR SER_MAIN+SER_CODE
zstr = "SELECT * FROM patoolslib.SERVICED WHERE SER_MAIN=? AND SER_CODE=? ORDER BY SER_MAIN,SER_CODE,SER_IND"
Set zsercodemap = rdb.CreateQuery("zsercmpdis", zstr)
zsercodemap(0) = "F"
zsercodemap(1) = "F1"
Set zrssercodemap = zsercodemap.OpenResultset(2, 1)


'Get the details from mapping table : SHI1.SERVICED FOR SER_MAIN+SER_IND
zstr = "SELECT * FROM patoolslib.SERVICED WHERE SER_MAIN=? AND SER_IND=? ORDER BY SER_MAIN,SER_CODE,SER_IND"
Set zserindmap = rdb.CreateQuery("zserimdis", zstr)
zserindmap(0) = "F"
zserindmap(1) = "F301"
Set zrsserindmap = zserindmap.OpenResultset(2, 1)

'****************************************************************************************************************************
 'This query is used to find out the status of the bed if the patient is
 'admitted and not given discharge intimation and not discharged for the
 'supplied case number
 Set Ma104_Qr = rdb.CreateQuery("IPma104Q", "select status from PATOOLSLIB.MA104 where case_no = ?") 'BILLOPT 61299 ADDED
' Ma104_Qr(0) = ">??/#####"
' Set ma104_rs = Ma104_Qr.OpenResultset(2, 3)
'****************************************************************************************************************************
 
 Set RSDBTPQR = rdb.CreateQuery("IPRSDBTPQ", "select ch_code from PATOOLSLIB.ma107 where  bed_type = ? AND  pat_cat = ?") 'BILLOPT 61299 ADDED
 
' RSDBTPQR(0) = "SP"
' RSDBTPQR(1) = "B"
'
' Set RSDBTP = RSDBTPQR.OpenResultset(2, 3) 'BILLOPT 61299 ADDED
 Set RSDNAMEQR = rdb.CreateQuery("IPRSDNAMEQ", "select * from PATOOLSLIB.ma102 where doc_code = ?") 'BILLOPT 61299 ADDED FOR DOCNAME IN DBGRID
' RSDNAMEQR(0) = "00000" 'BILLOPT 61299 ADDED FOR REQUERY
' Set rs_dname = RSDNAMEQR.OpenResultset(2, 3) 'BILLOPT 61299 ADDED FOR REQUERY
 
 Set RS5QR = rdb.CreateQuery("IPRS5Q", "select * from PATOOLSLIB.MA101 where ser_code =? or ser_det = ?") 'BILLOPT 61299 ADDED
 
 Set RS_CR_QR = rdb.CreateQuery("IPRS_CR_Q", "select * from PATOOLSLIB.ta114 where case_no = ?") 'BILLOPT 61299 ADDED
' RS_CR_QR(0) = ">??/#####" 'BILLOPT 61299 ADDED
' Set rs_cr = RS_CR_QR.OpenResultset(2, 3) 'BILLOPT 61299 ADDED
 
 Set RS_TR_QR = rdb.CreateQuery("IPRS_TR_Q", "select * from PATOOLSLIB.ta113 where case_no = ?") 'BILLOPT 61299 ADDED
' RS_TR_QR(0) = ">??/#####" 'BILLOPT 61299 ADDED
' Set rs_tr = RS_TR_QR.OpenResultset(2, 3) 'BILLOPT 61299 ADDED
 '****************
'' SQL = "INSERT INTO PATOOLSLIB.TA107(MEMO_NO,IO_MEMO,MEMO_DATE,DEPT_CODE,CASE_NO,MEMO_CATG,D_S_NAME,REM,DOC_CODE,AMT_RCVD,MEMO_AMT) VALUES (?,?,?,?,?,?,?,?,?,?,?)"
'''SQL = "INSERT INTO PATOOLSLIB.TA107(MEMO_NO,IO_MEMO,MEMO_DATE,DEPT_CODE,CASE_NO,memo_CATG,D_S_NAME,REM,REQ_NO,REQ_DATE,DOC_CODE,AMT_RCVD,MEMO_AMT) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"  'BILLOPT 61299 ADDED FOR INSERT INTO TA107
'' Set addta107_QR = rdb.CreateQuery("IPaddta107_q", SQL) 'BILLOPT 61299 ADDED FOR INSERT INTO TA107
''
'' SQL = "INSERT INTO PATOOLSLIB.TA1071 (MEMO_NO, SER_CODE, QTY, SER_AMT, DOC_CODE,REQ_NO,req_date,req_catg) VALUES (?, ?,?, ?,?,?, ?, ?)"
'' 'SQL = "INSERT INTO PATOOLSLIB.TA1071(MEMO_NO,SER_CODE,QTY,SER_AMT,DOC_CODE) VALUES (?,?,?,?,?)" 'BILLOPT 61299 ADDED FOR INSERT INTO TA1071
'' Set addta1071_QR = rdb.CreateQuery("IPaddta1071_q", SQL) 'BILLOPT 61299 ADDED FOR INSERT INTO TA1071
 
 '******
 
  sql = "INSERT INTO PATOOLSLIB.TA107(MEMO_NO,IO_MEMO,MEMO_DATE,DEPT_CODE,CASE_NO,MEMO_CATG,D_S_NAME,REM,DOC_CODE,AMT_RCVD,MEMO_AMT,SHIFLAG) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"
  Set addta107_QR = rdb.CreateQuery("IPaddta107_q", sql) 'BILLOPT 61299 ADDED FOR INSERT INTO TA107
 
 ''******change on 26/06/03
 
 sql = "INSERT INTO PATOOLSLIB.TA1071 (MEMO_NO, SER_CODE, QTY, SER_AMT, DOC_CODE,REQ_NO,req_date,req_catg,PROF_AMT,HOSP_AMT,DEPT_CODE,fld1) VALUES (?,?, ?,?, ?,?,?, ?, ?,?,?,?)"
 Set addta1071_QR = rdb.CreateQuery("IPaddta1071_q", sql) 'BILLOPT 61299 ADDED FOR INSERT INTO TA1071

 
 
 '******
 Set rs7qr = rdb.CreateQuery("IPRS7Q", "select * from PATOOLSLIB.TA1071 where memo_no = ?") 'BILLOPT 61299 ADDED FOR MEMO VIEW
'     rs7qr(0) = "M-000000" 'BILLOPT 61299 ADDED
'     Set rs7 = rs7qr.OpenResultset(2, 3)
     
'****************************************************************************************************************************
'This query counts the numerb if records having the supplied memo number
Set rs22qr = rdb.CreateQuery("IPrs22q", "select count(*) as c  from PATOOLSLIB.ta107 where memo_no = ?") 'billopt add
'****************************************************************************************************************************
      
 
 Set RS10QR = rdb.CreateQuery("IPRS10Q", "select * from PATOOLSLIB.TA1071 where memo_no = ?") 'BILLOPT 61299 ADDED FOR MEMO VIEW
'     RS10QR(0) = "M-000000" 'BILLOPT 61299 ADDED
'     Set rs10 = RS10QR.OpenResultset(2, 3)
 sql = "select * from PATOOLSLIB.TA107 where MEMO_NO = ? AND io_memo = ? "
Set RS44QR = rdb.CreateQuery("IPRS44Q", sql)

'****************************************************************************************************************************
'This query is used to check if the supplied memo number
'already exists in the database.
'NOTE: it does not distinguish between out-patient/In-patient
'memo numbers. That is the same number cannot be used.
sql = "SELECT * FROM PATOOLSLIB.TA107 WHERE MEMO_NO = ?"
Set Exist_ChkQr = rdb.CreateQuery("TATAHIA1", sql)
'Exist_ChkQr(0) = "M-100010"
'Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
'****************************************************************************************************************************


btn_vis (0)
Enblbox (0) '19599
'ADD ON 17072000
txt_req_no.MaxLength = 8
'THIS IS TAKEN FROM  DBGRID1_COLUMNCLICK TO FORM_LOAD 17072000
flex1.ZOrder
   Set RS = rdb.OpenResultset("select ser_code,ser_det from PATOOLSLIB.MA101", rdOpenKeyset, 3)
    While Not RS.EOF
     flex1.AddItem RS("ser_code") + vbTab + RS("ser_det")
     RS.MoveNext
    Wend
    flex2.ZOrder
    Set RS = rdb.OpenResultset("select doc_code from PATOOLSLIB.MA102", rdOpenKeyset, 3)
    While Not RS.EOF
     flex2.AddItem RS("doc_code")
     RS.MoveNext
    Wend
 'UNTIL HERE  FROM DBGRID1_COLUMNCLICK 17072000
'ADDED FOR RIS 11/05/2001
   'Set rq_reqdet = rdb.CreateQuery("rq_reqdet_id", "select reqnno,reqndt,patnstat,dept_code,pat_catg,ser_code,refbydoc from rrequisitn where caseno = ? and patnstat = ? and reqnno = ? and dept_code = ? and ((reqnstat is null) or reqnstat = 'A')")
   Set rq_reqdet = rdb.CreateQuery("rq_reqdet_id", "select reqnno,reqndt,patnstat,dept_code,pat_catg,ser_code,refbydoc from Rislib.rrequisitn where caseno = ? and patnstat = ? and reqnno = ?  and ((reqnstat is null) or reqnstat = 'A')")
   Set rq_updtreqdet = rdb.CreateQuery("rq_updtreqdet_id", "update Rislib.rrequisitn set memodt = ?,memono = ?,reqnstat = ? where caseno = ? and reqnno = ?  and ((reqnstat is null) or reqnstat = 'A')")
   'Set rq_updtreqdet = rdb.CreateQuery("rq_updtreqdet_id", "update rislib.rrequisitn set memodt = ?,memono = ?,reqnstat = ? where caseno = ? and reqnno = ? and dept_code = ? and ((reqnstat is null) or reqnstat = 'A')")
   Set rq_updtdelreqdet = rdb.CreateQuery("rq_updtdelreqdet_id", "update Rislib.rrequisitn set memono = ?,memodt = ?,reqnstat = ? where memono = ?")
  ' Set RQ_REQDETREQNO = rdb.CreateQuery("RQ_REQDETREQNO_ID", "SELECT reqnno FROM rrequisitn WHERE CASENO = ? AND patnstat = ? and dept_code = ? AND ((reqnstat is null)  or reqnstat = 'A')")
  ' if u want all departments requisitions remove the comment  below and add comments above chandu
    Set RQ_REQDETREQNO = rdb.CreateQuery("1RQ_REQDETREQNO_ID", "SELECT reqnno FROM Rislib.rrequisitn WHERE CASENO = ? AND patnstat = ?  AND ((reqnstat is null)  or reqnstat = 'A')")
 'ADDED FOR RIS 11/05/2001
  sql = "select pat_catg from Rislib.rrequisitn where caseno = ?"
  Set requisitn_qr = rdb.CreateQuery("requ_qr1", sql)
'  requisitn_qr(0) = "BN/00000"
'  Set requisitn_rs = requisitn_qr.OpenResultset(2, 3)
 
    Set REQUISITN_UPDT = rdb.CreateQuery("REQUISITN_UPDT1", " UPDATE RISLIB.RREQUISITN SET RECPTNO=?,RECPTDT=?,  REQNSTAT=? WHERE REQNNO=? and caseno=?")
  
  'SQL = "select req_no from rislib.ta1071 where memo_no=?"
  sql = "select req_no from PATOOLSLIB.ta1071 where memo_no=?"
  Set RQ_reqdet1 = rdb.CreateQuery("rq_reqdet11", sql)
'  RQ_reqdet1(0) = "M-00000"
'  Set rs_reqdet1 = RQ_reqdet1.OpenResultset(1, 1)
  Set UPDTA107_QR = rdb.CreateQuery("UPDTA107QR1", "UPDATE RISLIB.TA107 SET AMT_RCVD=? WHERE MEMO_NO=?")
   
  sql = "select case_no from RISLIB.ta105 where case_no = ? AND DISC_DATE IS NULL AND ADMS_DATE IS NOT NULL"
  Set InPt_ChkQr = rdb.CreateQuery("InPt_ChkQ", sql)
   
 me_date.Enabled = False
 
 sql = "SELECT YEAR1,SERIES FROM PATOOLSLIB.AUTO_MEMO WHERE YEAR1 =? "
 Set RQ_SERIES = rdb.CreateQuery("RQ_SERIES1", sql)
 
 'Set RS_SERIES = RQ_SERIES.OpenResultset(2, 3)
 
 'modified by v.s.k. on 13022012
 'Set rq_updtSERIES = rdb.CreateQuery("rq_updtSERIESmemosrno_id", "update PATOOLSLIB.AUTO_MEMO set SERIES = ? WHERE YEAR1 = ? ")  ' ADDED 11/11/2001
 Set rq_updtSERIES = rdb.CreateQuery("1rq_updtSERIESmemosrno_d", "update PATOOLSLIB.AUTO_MEMO set SERIES = ? WHERE YEAR1 = ? and SERIES<?")  ' ADDED 11/11/2001
 
 '***** SEKHAR IS ADDED THIS CODE FOR DIS
 sql = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
 Set SELDMULPROCEQR = rdb.CreateQuery("SELDMULPROCEQRIP", sql)
 
 sql = "SELECT * FROM DISLIB.DMOLEPROC WHERE REQNO=?"
 Set SELDMOLEPROCQR = rdb.CreateQuery("SELDMOLEPROCEQRIP", sql)
 
 sql = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
 Set zgetservice = rdb.CreateQuery("zgetservIPBM2", sql)
 
 '******
 'Get the details from mapping table : SHI1.SERVICED
sql = "SELECT * FROM PATOOLSLIB.SERVICED WHERE SER_MAIN=? ORDER BY SER_CODE"
Set zmaptable = rdb.CreateQuery("zmaptableq1IP", sql)
'zmaptable(0) = "A"
'Set zrsmaptable = zmaptable.OpenResultset(2, 1)

'Update Dept code and Doc Code as per the mapping table (Service Code - Doc Code and Dept code)
sql = "UPDATE PATOOLSLIB.TA1071 SET DEPT_CODE =?,DOC_CODE=? WHERE SER_CODE=? AND MEMO_NO=?"
Set zupdtmap = rdb.CreateQuery("zupdtmapq1IP", sql)
'zupdtmap(0) = "AAA"
'zupdtmap(1) = "99999"
'zupdtmap(2) = "A001"
'zupdtmap(3) = "M2999999"

'--------------------------------------------------------------------------------
'18/10/2002-Parag/Sandhya
'This query is used to get all records from SHI1.patoolslib_ta1071
sql = "SELECT * FROM PATOOLSLIB.TA1071 WHERE MEMO_NO=? ORDER BY SER_CODE"
Set zta1071 = rdb.CreateQuery("zta1071MAP1IP", sql)
'zta1071(0) = "M999999"
'Set zrsta1071 = zta1071.OpenResultset(2, 1)
'--------------------------------------------------------------------------------

 sql = "select * from rislib.rrequisitn where reqnNo=?"
 Set selreq_qr = rdb.CreateQuery("selreq_qrip", sql)
 

 '******
  End If
Dim REQNO As String

'' sandhya 1.8.2008 protocol patient starts

    str = "SELECT * FROM PATOOLSLIB.PROTOCASE WHERE CASE_NO = ? AND IRBNO = ? "
    Set PROTCASE_QR = rdb.CreateQuery("protCASE_QR1", str)
    
        
    str = "SELECT * FROM PATOOLSLIB.PROTSERVIC WHERE irbNO = ? and ser_code = ? and (CURRENT DATE BETWEEN FROMDT AND TODT) AND STATUS = 'C' "
    Set PROTSERV_QR = rdb.CreateQuery("PROTSERV_AQR1", str)
    
    
    str = "INSERT INTO PATOOLSLIB.PROTMEMO (MEMO_NO,MEMO_DATE,irbNO,CASE_NO,SER_code,SER_amt,prof_amt,hosp_amt,qty,MEMO_AMT,DEPT_CODE,DOC_CODE,PAT_TYPE,PAT_CATG,REQ_NO,req_date,IO_MEMO,AMT_RCVD,USERID,IPADDR) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set PROTMEMO_Qr = rdb.CreateQuery("PROTMEMO_tQ", str)
    
    str = "update PATOOLSLIB.PROTOCASE set ser_amt = ser_amt + ? where (case_no = ? AND IRBNO = ? )"
    Set PROTCASE_UPD = rdb.CreateQuery("PROTCASE_Upd", str)
    
    str = "select * from PATOOLSLIB.PROTMEMO where memo_no = ?"
    Set QR_PRTMEMO = rdb.CreateQuery("QRPRTMEMO3", str)

    'Update Dept code and Doc Code as per the mapping table (Service Code - Doc Code and Dept code)
    zstr = "UPDATE PATOOLSLIB.protmemo SET DEPT_CODE =?,DOC_CODE=? WHERE SER_CODE=? AND MEMO_NO=?"
    Set zupdprotmap = rdb.CreateQuery("zupdPRtmaP", zstr)
'' protocol patient ends
  
 ' **** SEKHAR IS ADDED THIS CODE FOR DIS
 
 '''' added by shraddha
'''' dated 16.8.2010 to check if service code exist in table ser_ref for duplication

str = "select * from patoolslib.ser_ref where SER_CODE=?"
Set qrser_ref = rdb.CreateQuery("", str)
'*************************
 
'''' ADDED BY SANDHYA DATED 7/7/2012 FOR BP RGJAY PATIENT

str = "select * from PATOOLSLIB.BPPATIENT where case_no = ?"
    Set QRBPPATN = rdb.CreateQuery("", str)

 str = "update PATOOLSLIB.BPPATIENT set ser_amt = ser_amt + ?, NET_AMT=SCHEM_AMT - ?  where (caseno = ?)"
    Set UPD_bppatn = rdb.CreateQuery("", str)
    
 str = "INSERT INTO PATOOLSLIB.BPPATMEMO (CASENO,PAT_CATG,MEMONO,MEMO_DT,MEMO_AMT,SCHEM_CODE,SCHEM_AMT,USERID,IPADD,FLAG)VALUES(?,?,?,?,?,?,?,?,?,?)"
    Set INSBPPATMEMO = rdb.CreateQuery("", str)
 
 ''''ADDED BY SANDHYA DATED 27-11-2012 TO UPDATE TRUST PATIENT'S BAL IN ACTREC-TMH
' Padmakar - (TMH-ACTREC) joint company account details
' ACTREC-TMH
zstr = "UPDATE PATOOLSLIB.TA113 SET NET_AMT=?, SER_AMT = ? WHERE  CASE_NO = ?"
Set s_UpdTa113 = rdb.CreateQuery("", zstr)

zstr = "SELECT MAX(MEMO_DATE) FROM PATOOLSLIB.TA107 WHERE CASE_NO=?"
Set selActMaxMemo = rdb.CreateQuery("", zstr)
 
 btn_vis (1)
End Sub



Private Sub DBgrid1_ButtonClick(ByVal ColIndex As Integer)
If ColIndex = 0 Then
' THIS IS ADDED TO FORM LOAD EVENT 17072000
    '   flex1.ZOrder
    '   Set RS = rdb.OpenResultset("select ser_code,ser_det from MA101", rdOpenKeyset, 3)
    '    While Not RS.EOF
    '     flex1.AddItem RS("ser_code") + vbTab + RS("ser_det")
    '     RS.MoveNext
    '    Wend
'UPTO HERE 17072000
ElseIf ColIndex = 4 And ncchg <> "Z" Then 'ADDED "AND NCCHG<> Z" 'ECTV
' THIS IS ADDED TO FORM LOAD EVENT 17072000
    '    flex2.ZOrder
    '    Set RS = rdb.OpenResultset("select doc_code from MA102", rdOpenKeyset, 3)
    '    While Not RS.EOF
    '     flex2.AddItem RS("doc_code")
    '     RS.MoveNext
    '    Wend
' UPTO HERE 17072000
End If
If a = 1 Then
  If ColIndex = 0 Then
    flex1.Visible = True
    flex1.SetFocus
    a = 0
  ElseIf ColIndex = 4 And ncchg <> "Z" Then  'ADDED "AND NCCHG<> Z" 'ECTV
    flex2.Visible = True
    flex2.SetFocus
    a = 0
  End If
ElseIf a = 0 Then
  If ColIndex = 0 Then
    flex1.Visible = False
    a = 1
  ElseIf ColIndex = 4 And ncchg <> "Z" Then  'ADDED "AND NCCHG<> Z" 'ECTV
    flex2.Visible = False
    a = 1
  End If
End If

If ColIndex = 0 Then
flex1.Top = DBGrid1.Top + DBGrid1.RowHeight * (DBGrid1.row + 2)
ElseIf ColIndex = 4 And ncchg <> "Z" Then  'ADDED "AND NCCHG<> Z" 'ECTV
flex2.Top = DBGrid1.Top + DBGrid1.RowHeight * (DBGrid1.row + 2)
End If
End Sub
Private Sub DBgrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
DBGrid1.Columns(4).Button = True
If ColIndex = 0 Then
flex1.Visible = False
ElseIf ColIndex = 4 Then
flex2.Visible = False
End If
a = 1

End Sub
Private Sub flex1_Click()
On Error GoTo MSG
'Dim opt As String
'Dim opttyp As String
'Dim ncchg As String '***22/10/98
If rs5flag = False Then
    RS5QR(0) = flex1.Text
    RS5QR(1) = flex1.Text
    Set rs5 = RS5QR.OpenResultset(1, 3)
    rs5flag = True
Else
    RS5QR(0) = flex1.Text
    RS5QR(1) = flex1.Text
    rs5.MoreResults
    rs5.Requery
End If
''' ADDED BY Sandhya protocol patient starts

PROTCASE_RS.Requery
        If PROTCASE_RS.RowCount = 1 Then
                 '  PROTSERV_QR(0) = PROTCASE_RS("IRBno")
                   PROTSERV_QR(0) = MIRBNO
                   PROTSERV_QR(1) = Left(DBGrid1.Columns(0), 4)
'                   PROTSERV_QR(2) = ZAS400DTTM!DT1
'                   PROTSERV_QR(3) = ZAS400DTTM!DT1
                   Set PROTSERV_RS = PROTSERV_QR.OpenResultset(2, 3)
                   If PROTSERV_RS.RowCount > 0 Then
                    Do While PROTSERV_RS.EOF
                      If RS_reqdet("ser_code") = PROTSERV_RS("ser_code") Then
                         DBGrid1.Columns(3) = PROTSERV_RS("ser_amt")
                         PTOT = PTOT + DBGrid1.Columns(3)
                         PROTFLG = True
                      Else
                      PROTSERV_RS.MoveNext
                      End If
                    
                    Loop
                   End If
          End If
'' protocol patient ends


''Set rs5 = rdb.OpenResultset("select * from MA101 where ser_code = '" & flex1.Text & "' or ser_det = '" & flex1.Text & "'", rdOpenKeyset, 3) 'BILLOPT 61299 COMMENTED
DBGrid1.Columns(0) = rs5("ser_code")
DBGrid1.Columns(1) = rs5("ser_det")
DBGrid1.Columns(2).Text = 1
DBgrid1_KeyPress (13)

'opt = lbl_cat
'opttyp = lbl_typ
'ncchg = Left(DBGrid1.Columns(0), 1) '***22/10/98
'EmptyGrid = 1 ''ADDED ON 17999
'Select Case opt
'
'Case "C "
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * rs5("C_CH")
'Case "B "
'   If opttyp <> "AF" Then '***22/10/98
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (rs5("B_HOSP_CH"))
'    Else
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (rs5("B_HOSP_CH") - rs5("B_PROF_CH"))  '***22/10/98
'    End If
'
'Case "A "
'    If opttyp <> "AF" Then '***22/10/98
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (rs5("A_HOSP_CH"))
'    Else
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (rs5("A_HOSP_CH") - rs5("A_PROF_CH")) '***22/10/98
'    End If
'
'Case "D "
'    If opttyp <> "AF" Then '***22/10/98
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (rs5("D_HOSP_CH"))
'    Else
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (rs5("D_HOSP_CH") - rs5("D_PROF_CH")) '***22/10/98
'    End If
'
'Case "F "
'    If opttyp <> "AF" Then '***22/10/98
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (rs5("FN_HOSP_CH"))
'    Else
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (rs5("FN_HOSP_CH") - rs5("FN_PROF_CH")) '***22/10/98
'    End If
'Case "NC"
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("NC_CH"))  '''''RRRR26799
''If ncchg <> "I" Then '**22/10/98
''    DBGrid1.Columns(3) = 0
''Else                 '**22/10/98
''    DBGrid1.Columns(3) = DBGrid1.Columns(2) * rs5("C_CH") '**22/10/98
''End If '**22/10/98
'
''    DBGrid1.Columns(3) = 0
'Case Else
'    MsgBox "NO MATCH ..."
'End Select
'
flex1.Visible = False
'DBGrid1.Col = 2
'DBGrid1.SetFocus
''DBGrid1.Row = DBGrid1.Row + 1
'A = 1
Exit Sub
MSG:  MsgBox Error(Err)

End Sub

Private Sub print1_Click()
'On Error Resume Next
'On Error GoTo ERROR_CODE:

'' COMMON PRINTING
'Dim DeviceName As String
'Dim CURPRT As String
'
'CURPRT = UCase(Printer.DeviceName)
''DeviceName = "Epson LQ-1050+"
'DeviceName = "COMMON"
'
'If DeviceName <> CURPRT Then
''If PRT_FLAG_STS = False Then
'   'XA = List1.List(List1.ListIndex)
'
'
'    If CSETPRINTER.SetPrinterAsDefault(DeviceName) Then
'        'sMsg = DeviceName & " has successfully been set as the default printer."
'     Else
'        MsgBox " " & DeviceName & " has failed to be set as the default printer."
'        Exit Sub
'    End If
''CommonDialog1.flags = &H40
''CommonDialog1.Action = 5
''End If
'End If
'
'CURPRT = UCase(Printer.DeviceName)
'
'If DeviceName <> CURPRT Then
'MsgBox "PLEASE CHANGE DEFULAT PRINTER TO " & DeviceName & " ", vbInformation
''CommonDialog1.flags = &H40
''CommonDialog1.Action = 5
'Exit Sub
'Else
''PRT_FLAG_STS = True
'End If
''COMMON PRINTING







If flag = 1 Then
    If rs2flag = False Then
        RS2QR(0) = CASENO
        Set rs2 = RS2QR.OpenResultset(1, 3)
        rs2flag = True
    Else
        RS2QR(0) = CASENO 'can be removed 'cause rs2 is already having required record
        rs2.MoreResults
        rs2.Requery
    End If
   ' MsgBox rs2.RowCount
    ''''Set rs2 = rdb.OpenResultset("select * from TA101 where case_no = '" & CASENO & "'", rdOpenKeyset)
    CR1.SelectionFormula = "{TA107.memo_no} = '" & "M-" & me_mno.Text & "'"    ' OR {TA107.memo_no} = '" & com_mno.Text & "' "
ElseIf flag = 2 Then
        If rs2flag = False Then
            RS2QR(0) = lbl_cno
            Set rs2 = RS2QR.OpenResultset(1, 3)
            rs2flag = True
        Else
            RS2QR(0) = lbl_cno 'can be removed 'cause rs2 is already having required record
            rs2.MoreResults
            rs2.Requery
        End If
        ''''Set rs2 = rdb.OpenResultset("select * from TA101 where case_no = '" & lbl_cno & "'", rdOpenKeyset)
        CR1.SelectionFormula = "{TA107.memo_no} = '" & "M-" & M_number & "'"
End If

'' SANDHYA PROTOCOL PATIENT STARTS
If PROTFLG = True And NPROTFLG = False Then
  PRTMEMOCR
ElseIf PROTFLG = True And NPROTFLG = True Then
  NPRTMEMOCR
ElseIf NPROTFLG = True Then
  ipmemocr
End If
'' PROTOCOL PATIENT ENDSipmemocr

'IPCR.CopiesToPrinter = 1
'IPCR.WindowState = crptMaximized
'IPCR.Destination = crptToWindow
'PrintResult% = IPCR.PrintReport
If PrintResult% = 0 Then
    Back.SetFocus
Exit Sub
Else
GoTo ERROR_CODE1
End If


'ipcr.Destination = crptToWindow
ipcr.Action = 1
  If rs2("PAT_TYPE") = "TA" Then
     CR1.Formulas(0) = "cr_tr = '" & rs_tr("TR_ACC_NO") & "'"
 ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = False Then
    CR1.Formulas(0) = "cr_tr = '" & Trim(rs_cr("CR_CARD")) & "'"
    CR1.Formulas(150) = "PATTP = '" & "COMPANY PATIENT" & "'"
   '''' ADDED BY SANDHYA DATED 18-7-2012 FOR RGJAY PATIENT
  If BPPATFLAG = True Then
   Set RS_BPPATN = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.BPPATIENT WHERE CASENO ='" & CASENO & "'", 1, 1)
  End If
   '''' ADDED BY SANDHYA DATED 7/7/2012 FOR RGJAY PATIENT
 ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = True Then
     CR1.Formulas(15) = "cr_tr = '" & Trim(RS_BPPATN("BPCARD_NO")) & "'"
     CR1.Formulas(150) = "PATTP = '" & "RGJAY PATIENT" & "'"
 ElseIf rs2("PAT_TYPE") = "CP" Then
    CR1.Formulas(0) = "cr_tr = '" & "    " & "'"
 End If
If flag = 1 Then
If IsDate(me_date.Text) Then
   If Format(me_date.Text, "dd/mm/yyyy") = Format(Now, "dd/mm/yyyy") Then
        'If memo date is current date
        If Label12.Caption = "UNIT NAME : " Then
            CR1.ReportFileName = "C:\PB121298\ip2.rpt"
        ElseIf Label12.Caption = "DOCTOR NAME : " Then
            CR1.ReportFileName = "C:\PB121298\ip.rpt"
        End If
   Else
        'If memo date is not a current date
        If Label12.Caption = "UNIT NAME : " Then
            CR1.ReportFileName = "C:\PB121298\Vip2.rpt"
        ElseIf Label12.Caption = "DOCTOR NAME : " Then
            CR1.ReportFileName = "C:\PB121298\Vip.rpt"
        End If
   End If
End If
ElseIf flag = 2 Then
If Label12.Caption = "UNIT NAME : " Then
     CR1.ReportFileName = "C:\PB121298\Vip2.rpt"
ElseIf Label12.Caption = "DOCTOR NAME : " Then
      CR1.ReportFileName = "C:\PB121298\Vip.rpt"
End If
End If

'CR1.Formulas(1) = "U_D_NAME = '" & txt_unit_doc_name & "'"
'CR1.Formulas(2) = "U_D_CODE = '" & lbl_dcode & "'"
'cr1.Formulas(1) = "UD_NAME = '" & txt_unit_doc_name & "'"
'cr1.Formulas(2) = "DOC_CODE = '" & lbl_dcode & "'"
'cr1.Action = 1
' Me.MousePointer = 11

'cr1.REPORT = crptToPrinter
'cr1.Action = 1
j = CR1.PrintReport
'Me.MousePointer = 0
Exit Sub
ERROR_CODE:
Dim er As rdoError
For Each er In rdoError
MsgBox er.Description
ERROR_CODE1:
If PrintResult% <> 0 Then
Exit Sub
End If
Next

On Error Resume Next
''''cr1.Formulas(0) = "cr_tr = " & " "
''''cr1.Formulas(1) = "U_D_NAME = " & " "
''''cr1.Formulas(2) = "U_D_CODE = " & " "
''''If flag = 1 Then
'''' Rs2Qr(0) = CASENO
'''' rs2.Requery
''''
''''cr1.SelectionFormula = "{TA107.memo_no} = '" & "M-" & me_mno.Text & "'"  ' OR {TA107.memo_no} = '" & com_mno.Text & "' "& "M-" & M_number.Text &
''''ElseIf flag = 2 Then
''''
''''Rs2Qr(0) = lbl_cno
''''rs2.Requery
''''
''''cr1.SelectionFormula = "{TA107.memo_no} = '" & "M-" & M_number.Text & "'" ' OR  '" & "M-" & M_number.Text & "' "
''''End If
''''
''''  If rs2("PAT_TYPE") = "TA" Then
''''     cr1.Formulas(0) = "cr_tr = '" & rs_tr("TR_ACC_NO") & "'"
''''   ElseIf rs2("PAT_TYPE") = "CR" Then
''''    cr1.Formulas(0) = "cr_tr = '" & rs_cr("CR_CARD") & "'"
''''   ElseIf rs2("PAT_TYPE") = "CP" Then
''''    cr1.Formulas(0) = "cr_tr = '" & "    " & "'"
''''   End If
''''
'''''If Label12.Caption = "UNIT NAME : " Then
'''''     cr1.ReportFileName = "C:\PB121298\ip2.rpt"
'''''ElseIf Label12.Caption = "DOCTOR NAME : " Then
'''''cr1.ReportFileName = "C:\PB121298\ip.rpt"
''''cr1.Formulas(1) = "U_D_NAME = '" & txt_unit_doc_name & "'"
''''cr1.Formulas(2) = "U_D_CODE = '" & lbl_dcode & "'"
''''    cr1.ReportFileName = "C:\PB121298\ip_dup.rpt"
'''''End If
''''cr1.Destination = crptToWindow
''''PrintResult% = cr1.PrintReport
''''
End Sub
Public Sub NPRTMEMOCR()
'' SANDHYA PROTOCOL PATIENT

Dim STR5 As String
Dim STR6 As String
Dim STR7 As String
Dim STR8 As String
Dim STR9 As String
Dim STR10 As String
Dim STR11 As String
Dim str12 As String
Dim CHKTOT As Double
Dim i As Integer
Dim j As Integer
Dim K As Integer
Dim samt As String
Dim l As Integer
Dim a As String
STR11 = COMP_IP
a = Chr(34)
samt = 0
i = 0
K = 0
ipcr.ReportFileName = "C:\PB121298\INPPRTMEMO.rpt"
ipcr.SelectionFormula = "{protmemo.MEMO_NO} = '" & me_mno.Text & "'"
'For J = 0 To 120 '261099
'    ipcr.Formulas(J) = " "
'Next J

If flag = 1 Then
'ipcr.Formulas(0) = "fmemono = '" & me_mno.Text & "'"    '"M-" & me_mno.Text & "'"
ipcr.Formulas(3) = "fdeptname = '" & Com_dno.Text & "'"
'ipcr.Formulas(2) = "fcaseno = '" & CASENO & "'"
str12 = CASENO
    If rs22_CHKflag = False Then
        Exist_ChkQr(0) = me_mno.Text   '"M-" & me_mno.Text
        Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
        rs22_CHKflag = True
    Else
        Exist_ChkQr(0) = me_mno.Text
        rs22_CHK.MoreResults
        rs22_CHK.Requery
    End If
STR10 = "*A*"
Else
'ipcr.Formulas(0) = "fmemono = '" & M_number & "'"     '"M-" & M_number & "'"
'ipcr.Formulas(2) = "fcaseno = '" & lbl_cno & "'"
ipcr.Formulas(3) = "fdeptname = '" & lbl_dname & "'"
     '"M-" & M_number
    str12 = lbl_cno
    If rs22_CHKflag = False Then
        Exist_ChkQr(0) = M_number
        Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
        rs22_CHKflag = True
    Else
        Exist_ChkQr(0) = M_number
        rs22_CHK.MoreResults
        rs22_CHK.Requery
    End If
STR10 = "*V*"
End If
 'If DOCFLAG = True Then
    ipcr.Formulas(1) = "doc_name = " & a & Trim(txt_unit_doc_name.Text) & a & ""
    ipcr.Formulas(8) = "doc_code = " & a & Trim(lbl_dcode.Caption) & a & ""
ipcr.Formulas(4) = "fdcode = " & a & lbl_dno & a & ""
 
ipcr.Formulas(5) = "NAME_CON = " & a & Trim(UCase(LBL_NAME)) & a & ""
ipcr.Formulas(10) = "ftot = " & a & " *****" & Format(lbl_tot_amt, "#,##,##,##,##0.00") & a & ""
ipcr.Formulas(6) = "TXTPRTtot = " & a & " *****" & Format(PTOT, "#,##,##,##,##0.00") & a & ""
ipcr.Formulas(7) = "TXTnet = " & a & " *****" & Format(TXTNET, "#,##,##,##,##0.00") & a & ""
ipcr.Formulas(12) = "fremark = " & a & Trim(txt_rem.Text) & a & ""
ipcr.Formulas(13) = "fmemodate  = " & a & Trim(me_date.Text) & a & ""
'STR1 = RTrim(username) 'change this var
 STR5 = RTrim(usr_name)
 STR6 = Format(DATE11, "DD/MM/YYYY")
 STR7 = time1
 'Exist_ChkQr(0) = "M-100010"
'Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
If rs22_CHK.RowCount <> 0 Then
 STR9 = Trim(rs22_CHK("D_S_NAME"))
End If
 'STR8 = Trim(STR5) + "  " + Trim(STR6) + "  " + Trim(STR7)
 STR8 = Trim(STR9) + " / " + Trim(STR5) + " " + Trim(STR10) + " " + " " + STR11 + " " + Trim(STR6) + "  " + Trim(STR7)
ipcr.Formulas(14) = "us_dt_tm = " & a & STR8 & a & ""

 If rs2("PAT_TYPE") = "TA" Then
     ipcr.Formulas(15) = "cr_tr = '" & rs_tr("TR_ACC_NO") & "'"
   ElseIf rs2("PAT_TYPE") = "CR" Then
    ipcr.Formulas(15) = "cr_tr = '" & rs_cr("CR_CARD") & "'"
   ElseIf PROTFLG = True Then
   ' ipcr.Formulas(9) = "cr_tr = '" & "    " & "'"
    'ipcr.Formulas(10) = "PATTP = '" & "PROTOCOL PATIENT" & "'"
   ElseIf rs2("PAT_TYPE") = "CP" Then
    ipcr.Formulas(15) = "cr_tr = '" & "    " & "'"
 End If
    

For Each p In Printers

ipcr.Destination = crptPRINTER
PrintResult% = ipcr.PrintReport
ipcr.Action = 1
If PrintResult% <> 0 Then
MsgBox Error(Err)
End If
Exit For
Next
End Sub

Public Sub PRTMEMOCR()
'' SANDHYA PROTOCOL PATIENT

Dim STR5 As String
Dim STR6 As String
Dim STR7 As String
Dim STR8 As String
Dim STR9 As String
Dim STR10 As String
Dim STR11 As String
Dim str12 As String
Dim CHKTOT As Double
Dim i As Integer
Dim j As Integer
Dim K As Integer
Dim samt As String
Dim l As Integer
Dim a As String
STR11 = COMP_IP
a = Chr(34)
samt = 0
i = 0
K = 0
ipcr.ReportFileName = "C:\PB121298\IPPRTMEMO.rpt"
ipcr.SelectionFormula = "{protmemo.MEMO_NO} = '" & me_mno.Text & "'"
'For J = 0 To 120 '261099
'    ipcr.Formulas(J) = " "
'Next J

If flag = 1 Then
'ipcr.Formulas(0) = "fmemono = '" & me_mno.Text & "'"    '"M-" & me_mno.Text & "'"
ipcr.Formulas(3) = "fdeptname = '" & Com_dno.Text & "'"
'ipcr.Formulas(2) = "fcaseno = '" & CASENO & "'"
str12 = CASENO
    If rs22_CHKflag = False Then
        Exist_ChkQr(0) = me_mno.Text   '"M-" & me_mno.Text
        Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
        rs22_CHKflag = True
    Else
        Exist_ChkQr(0) = me_mno.Text
        rs22_CHK.MoreResults
        rs22_CHK.Requery
    End If
STR10 = "*A*"
Else
'ipcr.Formulas(0) = "fmemono = '" & M_number & "'"     '"M-" & M_number & "'"
'ipcr.Formulas(2) = "fcaseno = '" & lbl_cno & "'"
ipcr.Formulas(3) = "fdeptname = '" & lbl_dname & "'"
     '"M-" & M_number
    str12 = lbl_cno
    If rs22_CHKflag = False Then
        Exist_ChkQr(0) = M_number
        Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
        rs22_CHKflag = True
    Else
        Exist_ChkQr(0) = M_number
        rs22_CHK.MoreResults
        rs22_CHK.Requery
    End If
STR10 = "*V*"
End If
 'If DOCFLAG = True Then
    ipcr.Formulas(1) = "doc_name = " & a & Trim(txt_unit_doc_name.Text) & a & ""
    ipcr.Formulas(8) = "doc_code = " & a & Trim(lbl_dcode.Caption) & a & ""
'ipcr.Formulas(4) = "fdcode = " & a & lbl_dno & a & ""
 
ipcr.Formulas(5) = "NAME_CON = " & a & Trim(UCase(LBL_NAME)) & a & ""
ipcr.Formulas(10) = "ftot = " & a & " *****" & Format(lbl_tot_amt, "#,##,##,##,##0.00") & a & ""
ipcr.Formulas(6) = "TXTPRTtot = " & a & " *****" & Format(PTOT, "#,##,##,##,##0.00") & a & ""
ipcr.Formulas(7) = "TXTnet = " & a & " *****" & Format(TXTNET, "#,##,##,##,##0.00") & a & ""
ipcr.Formulas(12) = "fremark = " & a & Trim(txt_rem.Text) & a & ""
'ipcr.Formulas(13) = "fmemodate  = " & a & Trim(me_date.Text) & a & ""
'STR1 = RTrim(username) 'change this var
 STR5 = RTrim(usr_name)
 STR6 = Format(DATE11, "DD/MM/YYYY")
 STR7 = time1
 'Exist_ChkQr(0) = "M-100010"
'Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
If rs22_CHK.RowCount <> 0 Then
 STR9 = Trim(rs22_CHK("D_S_NAME"))
End If
 'STR8 = Trim(STR5) + "  " + Trim(STR6) + "  " + Trim(STR7)
 STR8 = Trim(STR9) + " / " + Trim(STR5) + " " + Trim(STR10) + " " + " " + STR11 + " " + Trim(STR6) + "  " + Trim(STR7)
ipcr.Formulas(14) = "us_dt_tm = " & a & STR8 & a & ""

 If rs2("PAT_TYPE") = "TA" Then
     ipcr.Formulas(15) = "cr_tr = '" & rs_tr("TR_ACC_NO") & "'"
   ElseIf rs2("PAT_TYPE") = "CR" Then
    ipcr.Formulas(15) = "cr_tr = '" & rs_cr("CR_CARD") & "'"
   ElseIf PROTFLG = True Then
 '   ipcr.Formulas(9) = "cr_tr = '" & "    " & "'"
  '  ipcr.Formulas(10) = "PATTP = '" & "PROTOCOL PATIENT" & "'"
   ElseIf rs2("PAT_TYPE") = "CP" Then
    ipcr.Formulas(15) = "cr_tr = '" & "    " & "'"
 End If
    

For Each p In Printers

ipcr.Destination = crptPRINTER
PrintResult% = ipcr.PrintReport
ipcr.Action = 1
If PrintResult% <> 0 Then
MsgBox Error(Err)
End If
Exit For
Next

End Sub
Public Sub IPROTMAP()
'' SANDHYA PROTOCOL PATIENT
Dim zmatched As Boolean
Dim cd As String
zmatched = False
'Updating Doc_code and Dept_code against Service Code
'as per mapping table PATOOLSLIB/SERVICED.
Dim zPRTMEMOflag As Boolean
If zPRTMEMOflag = False Then
    QR_PRTMEMO(0) = me_mno.FormattedText
    Set RS_PRTMEMO = QR_PRTMEMO.OpenResultset(2, 1)
    RS_PRTMEMOflag = True
Else
    RS_PRTMEMO.MoreResults
    RS_PRTMEMO.Requery
End If
'############################################################################
'18/10/2002-Parag/Sandhya
'Updating Doc_code and Dept_code against Service Code
'as per mapping table PATOOLSLIB/SERVICED.
QR_PRTMEMO(0) = me_mno.FormattedText
RS_PRTMEMO.Requery

If RS_PRTMEMO.RowCount > 0 Then
RS_PRTMEMO.MoveFirst
DoEvents
    Do While RS_PRTMEMO.EOF = False
        zmatched = False
        
        If ZRSMAPTABLEFLAG = False Then
            zmaptable(0) = Left(Trim(RS_PRTMEMO!ser_code), 1)
            Set zrsmaptable = zmaptable.OpenResultset(2, 1)
            ZRSMAPTABLEFLAG = True
        Else
            zrsmaptable.MoreResults
            zrsmaptable.Requery
        End If
        
        zmaptable(0) = Left(Trim(RS_PRTMEMO!ser_code), 1)
        zrsmaptable.Requery
        If zrsmaptable.RowCount > 0 Then
           zrsmaptable.MoveFirst
           'case 1 :
           'check if the service code starts with "J"
     '      If Trim(zrsmaptable!SER_MAIN) = "J" Or Trim(zrsmaptable!SER_MAIN) = "O" Or Trim(zrsmaptable!SER_MAIN) = "Z" Or Trim(zrsmaptable!SER_MAIN) = "F" Or Trim(zrsmaptable!SER_MAIN) = "I" Or Trim(zrsmaptable!SER_MAIN) = "K" Or Trim(zrsmaptable!SER_MAIN) = "M" Then
                '--*---*--*---*--*---*--*---*--*---*--*---*--*---*--*---*--*---*
                 '28/03/2003-Parag/Sandhya
                 'For Those services where Doctor code is Not There.
                 'If (Trim(zrsmaptable!SER_IND)) = "" Then
                 Do While zrsmaptable.EOF = False
                    If ((Left(Trim(RS_PRTMEMO!ser_code), 1) = Trim(zrsmaptable!SER_MAIN)) And (Left(Trim(RS_PRTMEMO!ser_code), 2) = Trim(zrsmaptable!ser_code))) And Trim(RS_PRTMEMO!doc_code) = "000000" Then
                        zmatched = True
                        ZPART2 = False
                        zupdprotmap(0) = zrsmaptable!DEPT_CODE
                        zupdprotmap(1) = Trim(RS_PRTMEMO!doc_code)
                        zupdprotmap(2) = Trim(RS_PRTMEMO!ser_code)
                        zupdprotmap(3) = RS_PRTMEMO!memo_no
                        zupdprotmap.Execute
                        zmatched = True
                        Exit Do
                    End If
                     zrsmaptable.MoveNext
                 Loop
                 'End If
                 
                 '--*---*--*---*--*---*--*---*--*---*--*---*--*---*--*---*--*---*
            If zmatched = False Then 'start 1
               zrsmaptable.MoveFirst
            'For service code "J221","J222","J223".
             DoEvents
              Do While zrsmaptable.EOF = False
                 If Trim(RS_PRTMEMO!ser_code) = Trim(zrsmaptable!SER_IND) And Trim(RS_PRTMEMO!DEPT_CODE) = "X" Then
                    zupdprotmap(0) = zrsmaptable!DEPT_CODE
                    'zupdtmap(1) = IIf(Trim(zrsmaptable!doc_code) = "000000", Trim(zrsta1071!doc_code), Trim(zrsmaptable!doc_code))
                    zupdprotmap(1) = Trim(RS_PRTMEMO!doc_code)
                    zupdprotmap(2) = Trim(RS_PRTMEMO!ser_code)
                    zupdprotmap(3) = RS_PRTMEMO!memo_no
                    zupdprotmap.Execute
                    Exit Do
                 End If
                zrsmaptable.MoveNext
              Loop
            End If
            
              If zmatched = False Then
                 zrsmaptable.MoveFirst
                 'Check if Ser_code Value is There
                 If zrsmaptable!ser_code <> " " Then
                    zsercodemap(0) = zrsmaptable!SER_MAIN
                    zsercodemap(1) = Left(Trim(RS_PRTMEMO!ser_code), 2)
                    zrssercodemap.Requery
                    If zrssercodemap.RowCount > 0 Then
                        ZUSESERCODEMAP = True
                    Else
                        ZUSESERCODEMAP = False
                    End If
                 
                 End If
                 
                 'Check if Ser_ind Value Is there
                 If zrsmaptable!SER_IND <> " " Then
                    zserindmap(0) = zrsmaptable!SER_MAIN
                    zserindmap(1) = Trim(RS_PRTMEMO!ser_code)
                    zrsserindmap.Requery
                    If zrsserindmap.RowCount > 0 Then
                        ZUSESERINDMAP = True
                    Else
                        ZUSESERINDMAP = False
                    End If
                 End If
                 
                If ZUSESERINDMAP = True Then
                   zrsserindmap.MoveFirst
                   
                   
                   If (zrsserindmap!doc_code) = "000000" Then
                       zmatched = True
                       ZPART2 = False
                        zupdprotmap(0) = zrsserindmap!DEPT_CODE
                        zupdprotmap(1) = Trim(RS_PRTMEMO!doc_code)
                        zupdprotmap(2) = Trim(RS_PRTMEMO!ser_code)
                        zupdprotmap(3) = RS_PRTMEMO!memo_no
                        zupdprotmap.Execute
                   End If
                   
                   If zmatched = False Then
                   DoEvents
                   Do While zrsserindmap.EOF = False
                      If Trim(RS_PRTMEMO!doc_code) = Trim(zrsserindmap!doc_code) And Trim(RS_PRTMEMO!DEPT_CODE) = "X" Then
                            zmatched = True
                            ZPART2 = False
                            zupdprotmap(0) = zrsserindmap!DEPT_CODE
                            zupdprotmap(1) = Trim(zrsserindmap!doc_code)
                            zupdprotmap(2) = Trim(RS_PRTMEMO!ser_code)
                            zupdprotmap(3) = RS_PRTMEMO!memo_no
                            zupdprotmap.Execute
                            Exit Do
                      End If
                   zrsserindmap.MoveNext
                   Loop 'zrsserindmap.EOF = False
                   End If
                   
                   
                   If zmatched = False Then
                   zrsserindmap.MoveFirst
                     DoEvents
                     Do While zrsserindmap.EOF = False
                       If zrsserindmap!hod = "Y" Then
                        zmatched = True
                        ZPART2 = False
                        zupdprotmap(0) = zrsserindmap!DEPT_CODE
                        zupdprotmap(1) = Trim(zrsserindmap!doc_code)
                        zupdprotmap(2) = Trim(RS_PRTMEMO!ser_code)
                        zupdprotmap(3) = RS_PRTMEMO!memo_no
                        zupdprotmap.Execute
                        Exit Do
                       End If
                     zrsserindmap.MoveNext
                    Loop
                   End If
                End If
                
                
                If ZUSESERCODEMAP = True Then
                   zrssercodemap.MoveFirst
                   
                   If (zrssercodemap!doc_code) = "000000" Then
                       zmatched = True
                       ZPART2 = False
                        zupdprotmap(0) = zrssercodemap!DEPT_CODE
                        zupdprotmap(1) = Trim(RS_PRTMEMO!doc_code)
                        zupdprotmap(2) = Trim(RS_PRTMEMO!ser_code)
                        zupdprotmap(3) = RS_PRTMEMO!memo_no
                        zupdprotmap.Execute
                   End If
                   
                   DoEvents
                   If zmatched = False Then
                   Do While zrssercodemap.EOF = False
                      If Trim(RS_PRTMEMO!doc_code) = Trim(zrssercodemap!doc_code) And Trim(RS_PRTMEMO!DEPT_CODE) = "X" Then
                            zmatched = True
                            ZPART2 = False
                            zupdprotmap(0) = zrssercodemap!DEPT_CODE
                            zupdprotmap(1) = Trim(zrssercodemap!doc_code)
                            zupdprotmap(2) = Trim(RS_PRTMEMO!ser_code)
                            zupdprotmap(3) = RS_PRTMEMO!memo_no
                            zupdprotmap.Execute
                            Exit Do
                      End If
                   zrssercodemap.MoveNext
                   Loop 'zrsserindmap.EOF = False
                   End If
                   
                   If zmatched = False Then
                   zrssercodemap.MoveFirst
                     Do While zrssercodemap.EOF = False
                       If zrssercodemap!hod = "Y" Then
                        zmatched = True
                        ZPART2 = False
                        zupdprotmap(0) = zrssercodemap!DEPT_CODE
                        zupdprotmap(1) = Trim(zrssercodemap!doc_code)
                        zupdprotmap(2) = Trim(RS_PRTMEMO!ser_code)
                        zupdprotmap(3) = RS_PRTMEMO!memo_no
                        zupdprotmap.Execute
                        Exit Do
                       End If
                     zrssercodemap.MoveNext
                    Loop
                   End If
                End If 'zmatched = False
                 
                 
              If ZUSESERCODEMAP = False And ZUSESERINDMAP = False Then
                 zrsmaptable.Requery
                 zrsmaptable.MoveFirst
                 
                'if doctor code is "000000".
                If Trim(zrsmaptable!doc_code) = "000000" Then
                    zmatched = True
                    ZPART2 = False
                    zupdprotmap(0) = zrsmaptable!DEPT_CODE
                    zupdprotmap(1) = Trim(RS_PRTMEMO!doc_code)
                    zupdprotmap(2) = Trim(RS_PRTMEMO!ser_code)
                    zupdprotmap(3) = RS_PRTMEMO!memo_no
                    zupdprotmap.Execute
                End If
                
                DoEvents
                If zmatched = False Then
                   zrsmaptable.MoveFirst
                 Do While zrsmaptable.EOF = False
                    If Trim(RS_PRTMEMO!doc_code) = Trim(zrsmaptable!doc_code) And Trim(RS_PRTMEMO!DEPT_CODE) = "X" Then
                        zmatched = True
                        ZPART2 = False
                        zupdprotmap(0) = zrsmaptable!DEPT_CODE
                        zupdprotmap(1) = Trim(zrsmaptable!doc_code)
                        zupdprotmap(2) = Trim(RS_PRTMEMO!ser_code)
                        zupdprotmap(3) = RS_PRTMEMO!memo_no
                        zupdprotmap.Execute
                        Exit Do
                    End If
                 zrsmaptable.MoveNext
                 Loop
                 
                 'Assign HOD code as Doctor Code.
                 If zmatched = False Then
                 zrsmaptable.MoveFirst
                 Do While zrsmaptable.EOF = False
                   If zrsmaptable!hod = "Y" Then
                    zupdprotmap(0) = zrsmaptable!DEPT_CODE
                    zupdprotmap(1) = Trim(zrsmaptable!doc_code)
                    zupdprotmap(2) = Trim(RS_PRTMEMO!ser_code)
                    zupdprotmap(3) = RS_PRTMEMO!memo_no
                    zupdprotmap.Execute
                    Exit Do
                   End If
                 zrsmaptable.MoveNext
                Loop
                End If
                End If
              End If 'ZUSESERCODEMAP = False And ZUSESERINDMAP = False
             End If 'zmatched = False Then 'start 1
   End If
    RS_PRTMEMO.MoveNext
    ZPART2 = True
    ZUSESERCODEMAP = False
    ZUSESERINDMAP = False
    Loop
    
Else
MsgBox "No Records Available For Mapping", vbCritical, "Mapping Status"
Exit Sub
End If 'zrsta1071.RowCount > 0

End Sub
Private Sub txt_req_no_LOSTFocus()
If Len(Trim(txt_req_no.Text)) = 0 Then
txt_req_no.Text = "  "
End If

End Sub

Private Sub view_Click()
'******** copied form cleartxt function
lbl_cno.Caption = " "
'com_mno.Text = "  "
M_number = "  "
me_date.Mask = ""
me_date.Text = ""


ReDim UserData(0 To 5, 0 To 20)
DBGrid1.Refresh
flag = 2
'**********************
cleartxt


'Else
'While Not rs4.EOF
 'com_mno.AddItem rs4("memo_no")
'rs4.MoveNext
'Wend
'com_mno.Clear
d_u_name.Visible = False
commit.Enabled = False
cancel.Enabled = False
lbl_cno.Visible = True
me_cno.Visible = False
me_mno.Visible = False
'com_mno.Visible = True
M_number.Visible = True
'com_mno.Locked = False
lbl_dno.Visible = True
Com_dno.Visible = False
me_date.Enabled = False
'**********
'com_mno.Font = "ARIAL"
'com_mno.SetFocus
M_number.Enabled = True
'COMMENTED FOR AUTO GENERATION OF MEMO NO

'M_number.Mask = "M-######"
'COMMENTED FOR AUTO GENERATION OF MEMO NO
M_number.Mask = ">?#######"
'me_cno.Mask = ">??/#####"
me_cno.Text = "__/_____"
M_number.SelStart = 0

M_number.SetFocus

viewflag = True




'''@@@@############################333333333
'flag = 2
'Set rs4 = rdb.OpenResultset("select * from TA107 where io_memo = 'I'", rdOpenKeyset, 3)
'If rs4.RowCount = 0 Then
'  MsgBox "NO RECORDS ...."
'  New1.SetFocus
'Else
'While Not rs4.EOF
' com_mno.AddItem rs4("memo_no")
'rs4.MoveNext
'Wend
'End If
'commit.Enabled = False
'cancel.Enabled = False
'    lbl_cno.Caption = " "
'
'    com_mno.Text = "  "
'    me_date.Mask = ""
'    me_date.Text = ""
'    me_rdate.Mask = ""
'    me_rdate.Text = ""
'    ReDim UserData(0 To 5, 0 To 20)
'    DBGrid1.Refresh
''com_mno.Clear
'd_u_name.Visible = False
'lbl_cno.Visible = True
''Com_cno.Visible = False
'me_cno.Visible = False
'me_mno.Visible = False
'com_mno.Visible = True
'com_mno.Locked = False
'lbl_dno.Visible = True
'Com_dno.Visible = False
'me_date.Enabled = False
'CLEARTXT
'btn_vis (0)
'com_mno.SetFocus
'com_mno.Font = "ARIAL"
End Sub
Private Sub Refresh1()
'CLEARTXT
 M_number = " "
 lbl_cno = " "
 ReDim UserData(0 To 5, 0 To 20)
 DBGrid1.Refresh
 'New1.SetFocus
End Sub
Private Function TMH_ACTREC(CASENO As String)
' Padmakar - (TMH-ACTREC) joint company account details
         Call ACTREQMAIN
            If ACTtmhCn = True Then
                Call chkLastMemo(CASENO)
                If MaxMemo = True Then
                'Select memo date from ACTREC
                    selActMaxMemo(0) = CASENO
                    Set RsActMaxMemo = selActMaxMemo.OpenResultset(1, 1)
                    If RsActMaxMemo.RowCount <> 0 Then
                    ' find out lasted transations
                        If RsActMaxMemo(0) > maxdateACT Then
                            MaxMemo = True
                        Else
                            MaxMemo = False
                        End If
                    End If
            
                 'If last transaction  done in TMH then
                    If MaxMemo = False Then
                    Call ACTTMH_TABal(CASENO)
                        ' Update latest Total Service Amount (TMH + ACTREC)
                        ' Date - 09 Aug 2012
                        s_UpdTa113(0) = CRtOT
                        s_UpdTa113(1) = drTOT
                        s_UpdTa113(2) = CASENO
                        s_UpdTa113.Execute
                    Else
                    
                    End If
             End If
            End If
            '---------------------------------------------------------------------------
End Function
Sub VALIDATEAMT() ''19599
Dim tot As Long
Dim iRow As Integer
Text22 = " "
tot = 0
For iRow = 0 To mTotalRows&
    If Len(Trim(UserData(0, iRow))) > 0 Then '''''25041999
        If Not IsEmpty(UserData(0, iRow)) Then
               tot = tot + CLng(UserData(3, iRow))
        End If
    End If
Next iRow
''*****5599
Dim ttest As Long
If Len(Trim(lbl_tot_amt)) <> 0 Then
ttest = CDbl(lbl_tot_amt)
Else
ttest = 0
End If
If ttest <> tot Or ttest = 0 Then
AmtNotEq = 1
End If
''*****5599
End Sub
Sub GETSMARTCARD(case_num As String)
 '''' ADDED BY SANDHYA FOR SMART CARD PATIENT
'''' TO CHECK IF SMARTCARD PATIENT
    Dim RQ_SMARTCARD As rdoQuery
    Dim RS_SMARTCARD As rdoResultset
    
    Set RQ_SMARTCARD = rdb.CreateQuery("SMT", "SELECT * FROM PATOOLSLIB.SCARDACC WHERE CASE_NO=? AND AIFLAG <> 'I' ")
    RQ_SMARTCARD(0) = CASENO
       
    Set RS_SMARTCARD = RQ_SMARTCARD.OpenResultset(1, 1)
    
    If RS_SMARTCARD.RowCount > 0 Then
        SMCARDFLG = True
        lbl_smartcard.Visible = True
        lbl_deposit.Visible = True
        lbl_deposit.Caption = scard_bal_amt
        If RS_SMARTCARD("aiflag") = "L" Then
           MsgBox ("SMARTCARD HAS BEEN LOCKED, ACTIVATE SMART CARD, PLEASE CONTACT BILLING DEPT. TO ACTIVATE CARD")
        ElseIf RS_SMARTCARD("AIFLAG") = "A" Then
           MsgBox ("SMARTCARD SERVICE WILL BE DEDUCTED FROM ACCOUNT")
        End If
        scard_bal_amt = RS_SMARTCARD("bal_amt")
        
        
        
        ''''ADDED FOR WRITING IN ERROR LOG
        
        ''''UPTO HERE

        
        
        
         
      '  GETSMARTCARD = 1
        
'        DisplayPole ("Case No.: " & CASENO & "  " & "Balance " & PadSpace(12, RS_SMARTCARD("bal_amt")))
    Else
     lbl_smartcard.Visible = False
     lbl_deposit.Visible = False
     scard = False
    
     '   GETSMARTCARD = 0
    End If
    
    
    'MsgBox ("Please place the card on the Reader")
    RQ_SMARTCARD.Close
    Set RQ_SMARTCARD = Nothing
    Set RS_SMARTCARD = Nothing
'''' END OF SMART CARD
End Sub
Public Sub INSERTERROR(MEMONO As String, CASENO As String, CARDBAL As String, SYSBAL As String)
'''' ADDED BY SANDHYA FOR SMART CARD PATIENT
    Dim RQ_INSERTERROR As rdoQuery
    Dim D1 As String
    D1 = Format(me_date.Text, "dd/mm/yyyy")
    D1 = GET_REVDATE(D1)
    
    Set RQ_INSERTERROR = rdb.CreateQuery("INS", "INSERT INTO PATOOLSLIB.SERRORLOG(MEMO_NO,CASE_NO,CARDBALAMT,SYSBALAMT,USERID,IPADDRESS) VALUES(?,?,?,?,?,?)")
    RQ_INSERTERROR(0) = MEMONO
    RQ_INSERTERROR(1) = CASENO
    RQ_INSERTERROR(2) = Val(CARDBAL)
    RQ_INSERTERROR(3) = SYSBAL
        
    If Len(Trim(txt_doc_sec_name)) = 0 Then
        RQ_INSERTERROR(4) = "userid"
    Else
        RQ_INSERTERROR(4) = txt_doc_sec_name
    End If
    RQ_INSERTERROR(5) = Winsock1.LocalIP
    RQ_INSERTERROR.Execute
    
    
    Set RQ_INSERTERROR = Nothing
 '''' END OF SMART CARD
End Sub
Public Sub Enblbox(num) '''''5599
If num = 0 Then
 txt_unit_doc_name.Enabled = False
 txt_doc_sec_name.Enabled = False
 Com_dno.Enabled = False
 d_u_name.Enabled = False
 txt_rem.Enabled = False
 me_mno.Enabled = False
 me_cno.Enabled = False
 'me_date.Enabled = False
 me_date.Enabled = True
 DBGrid1.Enabled = False
 txt_req_no.Enabled = False
 me_rdate.Enabled = False
 ElseIf num = 1 Then
 txt_unit_doc_name.Enabled = True
 txt_doc_sec_name.Enabled = True
 Com_dno.Enabled = True
 d_u_name.Enabled = True
 txt_rem.Enabled = True
 me_mno.Enabled = True
 me_cno.Enabled = True
 me_date.Enabled = True
 DBGrid1.Enabled = True
 txt_req_no.Enabled = True
 me_rdate.Enabled = True
 End If
 End Sub


Private Sub NEW_LVL_BTN()
Dim RS_LVL As rdoResultset
Dim RS_FM As rdoResultset

'Set RS_LVL = rdb.OpenResultset("select * from TASEC2 where TRIM(USR_GRP) = '" & Trim(USRGRP) & "' AND TRIM(SCREEN) = '" & UCase(Trim(reg_doc.Name)) & "'", 2, 3)
Set RS_LVL = rdb.OpenResultset("select * from TOOLSADMN.scrn_auth1 where TRIM(USR_GRP) = '" & Trim(USRGRP) & "' AND TRIM(SCRN_ID) = '" & UCase(Trim(Me.name)) & "'", 2, 3)
'**14/6/99
LVL_AUTH = 1
If RS_LVL.RowCount = 0 Then
MsgBox "USER NOT AUTHORISED TO ACCESS THE SCREEN"
LVL_AUTH = 0
Exit Sub
End If
'**14/6/99
LVL = RS_LVL("LEVEL")
If Trim(RS_LVL("NEW_ENBL")) = "Y" Then
 New1.Enabled = True
ElseIf Trim(RS_LVL("NEW_ENBL")) = "N" Then
 New1.Enabled = False
End If


If Trim(RS_LVL("MOD_ENBL")) = "Y" Then
 'update.Enabled = True
ElseIf Trim(RS_LVL("MOD_ENBL")) = "N" Then
 'update.Enabled = False
End If

If Trim(RS_LVL("VIEW_ENBL")) = "Y" Then
 View.Enabled = True
ElseIf Trim(RS_LVL("VIEW_ENBL")) = "N" Then
 View.Enabled = False
End If

If Trim(RS_LVL("DEL_ENBL")) = "Y" Then
 Delete.Enabled = True
ElseIf Trim(RS_LVL("DEL_ENBL")) = "N" Then
 Delete.Enabled = False
End If

If Trim(RS_LVL("PRINT_ENBL")) = "Y" Then
 print1.Enabled = True
ElseIf Trim(RS_LVL("PRINT_ENBL")) = "N" Then
 print1.Enabled = False
End If

exit1.Enabled = True
Back.Enabled = True

End Sub

Public Sub CHKDOCCODE()
If rs18flag = False Then
    RS18QR(0) = DBGrid1.Columns(0)
    Set RS18 = RS18QR.OpenResultset(1, 3)
    rs18flag = True
Else
    RS18QR(0) = DBGrid1.Columns(0)
    RS18.MoreResults
    RS18.Requery
End If
'Set RS18 = rdb.OpenResultset("select * from MA101 where ser_code = '" & DBGrid1.Columns(0) & "'", rdOpenKeyset, 3) 'BILLOPT 61299 COMMENTED
opt = lbl_cat

Select Case opt

Case "B "
   If RS18("B_PROF_CH") <> 0 Then
      If Val(DBGrid1.Text) = 0 Then
        MsgBox "ENTER DOCTOR CODE....."
        DBGrid1.col = 4
        DBGrid1.SelLength = 5
        DBGrid1.SetFocus
      End If
   Else
       DBGrid1.Text = " " ' ADDED TO MAKE THE DOC CODE BLANK WHEN B_PROF_CH = 0
       DBGrid1.row = DBGrid1.row + 1
       DBGrid1.col = 0
   End If
  
Case "A "
     If RS18("A_PROF_CH") <> 0 Then
     If Val(DBGrid1.Text) = 0 Then
        MsgBox "ENTER DOCTOR CODE....."
        DBGrid1.col = 4
        DBGrid1.SelLength = 5
        DBGrid1.SetFocus
      End If
   Else
       DBGrid1.Text = " " ' ADDED TO MAKE THE DOC CODE BLANK WHEN A_PROF_CH = 0
       DBGrid1.row = DBGrid1.row + 1
       DBGrid1.col = 0
   End If
    
Case "D "
    If RS18("D_PROF_CH") <> 0 Then
    If Val(DBGrid1.Text) = 0 Then
        MsgBox "ENTER DOCTOR CODE....."
        DBGrid1.col = 4
        DBGrid1.SelLength = 5
        DBGrid1.SetFocus
     End If
   Else
       DBGrid1.Text = " " ' ADDED TO MAKE THE DOC CODE BLANK WHEN D_PROF_CH = 0
       DBGrid1.row = DBGrid1.row + 1
       DBGrid1.col = 0
   End If
'Case "FN"
Case "F "
    If RS18("FN_PROF_CH") <> 0 Then
    If Val(DBGrid1.Text) = 0 Then
        MsgBox "ENTER DOCTOR CODE....."
        DBGrid1.col = 4
        DBGrid1.SelLength = 5
        DBGrid1.SetFocus
    End If
   Else
       DBGrid1.Text = " " ' ADDED TO MAKE THE DOC CODE BLANK WHEN FN_PROF_CH = 0
       DBGrid1.row = DBGrid1.row + 1
       DBGrid1.col = 0
   End If

Case Else
    DBGrid1.row = DBGrid1.row + 1
    DBGrid1.col = 0
End Select
End Sub
Function MatchUp(txtusername As String)
    MatchUp = Switch(txtusername = "SANIL", True, txtusername = "COELHO", True, txtusername = "PRABHU", True, txtusername = "AMITA", True, txtusername = "SUMANT", True, txtusername = "LAXMAN", True, txtusername = "USHA", True)
End Function


Public Sub MAPING()
Dim zmatched As Boolean
Dim cd As String
zmatched = False
'Updating Doc_code and Dept_code against Service Code
'as per mapping table PATOOLSLIB/SERVICED.
If Zrsta1071flag = False Then
    zta1071(0) = me_mno.FormattedText
    Set zrsta1071 = zta1071.OpenResultset(2, 1)
    Zrsta1071flag = True
Else
    zrsta1071.MoreResults
    zrsta1071.Requery
End If
'############################################################################
'18/10/2002-Parag/Sandhya
'Updating Doc_code and Dept_code against Service Code
'as per mapping table PATOOLSLIB/SERVICED.
DoEvents
zta1071(0) = me_mno.FormattedText
zrsta1071.Requery
If zrsta1071.RowCount > 0 Then
zrsta1071.MoveFirst
    Do While zrsta1071.EOF = False
        zmatched = False
        
        If ZRSMAPTABLEFLAG = False Then
            zmaptable(0) = Left(Trim(zrsta1071!ser_code), 1)
            Set zrsmaptable = zmaptable.OpenResultset(2, 1)
            ZRSMAPTABLEFLAG = True
        Else
            zrsmaptable.MoreResults
            zrsmaptable.Requery
        End If
        
        zmaptable(0) = Left(Trim(zrsta1071!ser_code), 1)
        zrsmaptable.Requery
        If zrsmaptable.RowCount > 0 Then
           zrsmaptable.MoveFirst
           'case 1 :
           'check if the service code starts with "J"
     '      If Trim(zrsmaptable!SER_MAIN) = "J" Or Trim(zrsmaptable!SER_MAIN) = "O" Or Trim(zrsmaptable!SER_MAIN) = "Z" Or Trim(zrsmaptable!SER_MAIN) = "F" Or Trim(zrsmaptable!SER_MAIN) = "I" Or Trim(zrsmaptable!SER_MAIN) = "K" Or Trim(zrsmaptable!SER_MAIN) = "M" Then
                '--*---*--*---*--*---*--*---*--*---*--*---*--*---*--*---*--*---*
                 '28/03/2003-Parag/Sandhya
                 'For Those services where Doctor code is Not There.
                 'If (Trim(zrsmaptable!SER_IND)) = "" Then
                 DoEvents
                 Do While zrsmaptable.EOF = False
                    If ((Left(Trim(zrsta1071!ser_code), 1) = Trim(zrsmaptable!SER_MAIN)) And (Left(Trim(zrsta1071!ser_code), 2) = Trim(zrsmaptable!ser_code))) And Trim(zrsta1071!doc_code) = "000000" Then
                        zmatched = True
                        ZPART2 = False
                        zupdtmap(0) = zrsmaptable!DEPT_CODE
                        zupdtmap(1) = Trim(zrsta1071!doc_code)
                        zupdtmap(2) = Trim(zrsta1071!ser_code)
                        zupdtmap(3) = zrsta1071!memo_no
                        zupdtmap.Execute
                        zmatched = True
                        Exit Do
                    End If
                     zrsmaptable.MoveNext
                 Loop
                 'End If
                 
                 '--*---*--*---*--*---*--*---*--*---*--*---*--*---*--*---*--*---*
            If zmatched = False Then 'start 1
               zrsmaptable.MoveFirst
            'For service code "J221","J222","J223".
              DoEvents
              Do While zrsmaptable.EOF = False
                 If Trim(zrsta1071!ser_code) = Trim(zrsmaptable!SER_IND) And Trim(zrsta1071!DEPT_CODE) = "X" Then
                    zupdtmap(0) = zrsmaptable!DEPT_CODE
                    'zupdtmap(1) = IIf(Trim(zrsmaptable!doc_code) = "000000", Trim(zrsta1071!doc_code), Trim(zrsmaptable!doc_code))
                    zupdtmap(1) = Trim(zrsta1071!doc_code)
                    zupdtmap(2) = Trim(zrsta1071!ser_code)
                    zupdtmap(3) = zrsta1071!memo_no
                    zupdtmap.Execute
                    Exit Do
                 End If
                zrsmaptable.MoveNext
              Loop
            End If
            
              If zmatched = False Then
                 zrsmaptable.MoveFirst
                 'Check if Ser_code Value is There
                 If zrsmaptable!ser_code <> " " Then
                    zsercodemap(0) = zrsmaptable!SER_MAIN
                    zsercodemap(1) = Left(Trim(zrsta1071!ser_code), 2)
                    zrssercodemap.Requery
                    If zrssercodemap.RowCount > 0 Then
                        ZUSESERCODEMAP = True
                    Else
                        ZUSESERCODEMAP = False
                    End If
                 
                 End If
                 
                 'Check if Ser_ind Value Is there
                 If zrsmaptable!SER_IND <> " " Then
                    zserindmap(0) = zrsmaptable!SER_MAIN
                    zserindmap(1) = Trim(zrsta1071!ser_code)
                    zrsserindmap.Requery
                    If zrsserindmap.RowCount > 0 Then
                        ZUSESERINDMAP = True
                    Else
                        ZUSESERINDMAP = False
                    End If
                 End If
                 
                If ZUSESERINDMAP = True Then
                   zrsserindmap.MoveFirst
                   
                   
                   If (zrsserindmap!doc_code) = "000000" Then
                       zmatched = True
                       ZPART2 = False
                        zupdtmap(0) = zrsserindmap!DEPT_CODE
                        zupdtmap(1) = Trim(zrsta1071!doc_code)
                        zupdtmap(2) = Trim(zrsta1071!ser_code)
                        zupdtmap(3) = zrsta1071!memo_no
                        zupdtmap.Execute
                   End If
                   
                   If zmatched = False Then
                   DoEvents
                   Do While zrsserindmap.EOF = False
                      If Trim(zrsta1071!doc_code) = Trim(zrsserindmap!doc_code) And Trim(zrsta1071!DEPT_CODE) = "X" Then
                            zmatched = True
                            ZPART2 = False
                            zupdtmap(0) = zrsserindmap!DEPT_CODE
                            zupdtmap(1) = Trim(zrsserindmap!doc_code)
                            zupdtmap(2) = Trim(zrsta1071!ser_code)
                            zupdtmap(3) = zrsta1071!memo_no
                            zupdtmap.Execute
                            Exit Do
                      End If
                   zrsserindmap.MoveNext
                   Loop 'zrsserindmap.EOF = False
                   End If
                   
                   
                   If zmatched = False Then
                   zrsserindmap.MoveFirst
                     DoEvents
                     Do While zrsserindmap.EOF = False
                       If zrsserindmap!hod = "Y" Then
                        zmatched = True
                        ZPART2 = False
                        zupdtmap(0) = zrsserindmap!DEPT_CODE
                        zupdtmap(1) = Trim(zrsserindmap!doc_code)
                        zupdtmap(2) = Trim(zrsta1071!ser_code)
                        zupdtmap(3) = zrsta1071!memo_no
                        zupdtmap.Execute
                        Exit Do
                       End If
                     zrsserindmap.MoveNext
                    Loop
                   End If
                End If
                
                
                If ZUSESERCODEMAP = True Then
                   zrssercodemap.MoveFirst
                   
                   If (zrssercodemap!doc_code) = "000000" Then
                       zmatched = True
                       ZPART2 = False
                        zupdtmap(0) = zrssercodemap!DEPT_CODE
                        zupdtmap(1) = Trim(zrsta1071!doc_code)
                        zupdtmap(2) = Trim(zrsta1071!ser_code)
                        zupdtmap(3) = zrsta1071!memo_no
                        zupdtmap.Execute
                   End If
                   
                   
                   If zmatched = False Then
                   DoEvents
                   Do While zrssercodemap.EOF = False
                      If Trim(zrsta1071!doc_code) = Trim(zrssercodemap!doc_code) And Trim(zrsta1071!DEPT_CODE) = "X" Then
                            zmatched = True
                            ZPART2 = False
                            zupdtmap(0) = zrssercodemap!DEPT_CODE
                            zupdtmap(1) = Trim(zrssercodemap!doc_code)
                            zupdtmap(2) = Trim(zrsta1071!ser_code)
                            zupdtmap(3) = zrsta1071!memo_no
                            zupdtmap.Execute
                            Exit Do
                      End If
                   zrssercodemap.MoveNext
                   Loop 'zrsserindmap.EOF = False
                   End If
                   
                   If zmatched = False Then
                   zrssercodemap.MoveFirst
                     DoEvents
                     Do While zrssercodemap.EOF = False
                       If zrssercodemap!hod = "Y" Then
                        zmatched = True
                        ZPART2 = False
                        zupdtmap(0) = zrssercodemap!DEPT_CODE
                        zupdtmap(1) = Trim(zrssercodemap!doc_code)
                        zupdtmap(2) = Trim(zrsta1071!ser_code)
                        zupdtmap(3) = zrsta1071!memo_no
                        zupdtmap.Execute
                        Exit Do
                       End If
                     zrssercodemap.MoveNext
                    Loop
                   End If
                End If 'zmatched = False
                 
                 
              If ZUSESERCODEMAP = False And ZUSESERINDMAP = False Then
                 zrsmaptable.Requery
                 zrsmaptable.MoveFirst
                 
                'if doctor code is "000000".
                If Trim(zrsmaptable!doc_code) = "000000" Then
                    zmatched = True
                     ZPART2 = False
                    zupdtmap(0) = zrsmaptable!DEPT_CODE
                    zupdtmap(1) = Trim(zrsta1071!doc_code)
                    zupdtmap(2) = Trim(zrsta1071!ser_code)
                    zupdtmap(3) = zrsta1071!memo_no
                    zupdtmap.Execute
                End If
                
                
                If zmatched = False Then
                   zrsmaptable.MoveFirst
                 DoEvents
                 Do While zrsmaptable.EOF = False
                    If Trim(zrsta1071!doc_code) = Trim(zrsmaptable!doc_code) And Trim(zrsta1071!DEPT_CODE) = "X" Then
                        zmatched = True
                        ZPART2 = False
                        zupdtmap(0) = zrsmaptable!DEPT_CODE
                        zupdtmap(1) = Trim(zrsmaptable!doc_code)
                        zupdtmap(2) = Trim(zrsta1071!ser_code)
                        zupdtmap(3) = zrsta1071!memo_no
                        zupdtmap.Execute
                        Exit Do
                    End If
                 zrsmaptable.MoveNext
                 Loop
                 
                 'Assign HOD code as Doctor Code.
                 If zmatched = False Then
                 DoEvents
                 zrsmaptable.MoveFirst
                 Do While zrsmaptable.EOF = False
                   If zrsmaptable!hod = "Y" Then
                    zupdtmap(0) = zrsmaptable!DEPT_CODE
                    zupdtmap(1) = Trim(zrsmaptable!doc_code)
                    zupdtmap(2) = Trim(zrsta1071!ser_code)
                    zupdtmap(3) = zrsta1071!memo_no
                    zupdtmap.Execute
                    Exit Do
                   End If
                 zrsmaptable.MoveNext
                Loop
                End If
                End If
              End If 'ZUSESERCODEMAP = False And ZUSESERINDMAP = False
             End If 'zmatched = False Then 'start 1
   End If
    zrsta1071.MoveNext
    ZPART2 = True
    ZUSESERCODEMAP = False
    ZUSESERINDMAP = False
    Loop
Else
MsgBox "No Records Available For Mapping", vbCritical, "Mapping Status"
Exit Sub
End If 'zrsta1071.RowCount > 0

End Sub
