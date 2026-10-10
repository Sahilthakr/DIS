VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form op_memo_d1 
   BorderStyle     =   0  'None
   Caption         =   "op memo"
   ClientHeight    =   8400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8880
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox MSComm1 
      Height          =   480
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   68
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox nettot 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7800
      TabIndex        =   63
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox prottot 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   61
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   8280
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport ipcr 
      Left            =   3600
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   1692
      Left            =   120
      OleObjectBlob   =   "op_memo_d1.frx":0000
      TabIndex        =   54
      Top             =   3720
      Width           =   9132
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   6840
      Width           =   735
   End
   Begin VB.TextBox txt_doc_sec_name 
      Height          =   285
      Left            =   7200
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6360
      Width           =   1575
   End
   Begin VB.ComboBox Com_dno 
      Height          =   288
      Left            =   5880
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   3372
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
      Left            =   3600
      TabIndex        =   33
      Top             =   6840
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6840
      Width           =   1095
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
      Left            =   7080
      TabIndex        =   38
      Top             =   6840
      Width           =   1095
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
      Height          =   315
      Left            =   1800
      TabIndex        =   32
      Top             =   6885
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox txt_rem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   288
      Left            =   1080
      TabIndex        =   23
      Top             =   6360
      Width           =   3132
   End
   Begin VB.TextBox txt_unit_doc_name 
      Height          =   285
      Left            =   5400
      TabIndex        =   11
      Text            =   " "
      Top             =   2880
      Width           =   3615
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
      TabIndex        =   49
      Top             =   7440
      Width           =   5895
   End
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6840
      Width           =   1095
   End
   Begin MSMask.MaskEdBox me_cno 
      Height          =   372
      Left            =   4320
      TabIndex        =   6
      Top             =   1200
      Width           =   1452
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16711680
      PromptInclude   =   0   'False
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
      Format          =   ";;000000;;"
      Mask            =   ">??/#####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox me_mno 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      BackColor       =   0
      ForeColor       =   16777215
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      Enabled         =   0   'False
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
      Format          =   "M-######"
      Mask            =   ">??######"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox me_date 
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   128
      ForeColor       =   -2147483634
      AutoTab         =   -1  'True
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
   Begin MSFlexGridLib.MSFlexGrid flex1 
      Height          =   1692
      Left            =   120
      TabIndex        =   35
      Top             =   3720
      Width           =   7812
      _ExtentX        =   13785
      _ExtentY        =   2990
      _Version        =   393216
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
      GridLines       =   3
      SelectionMode   =   1
      AllowUserResizing=   3
      MousePointer    =   4
      FormatString    =   "^          CODE            |                          SERVICES/PROCEDURES/INVESTIGATIONS"
   End
   Begin MSFlexGridLib.MSFlexGrid flex2 
      Height          =   1692
      Left            =   4080
      TabIndex        =   36
      Top             =   3720
      Width           =   1692
      _ExtentX        =   2990
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
   Begin VB.ComboBox d_u_name 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2880
      Width           =   3852
   End
   Begin MSMask.MaskEdBox M_number 
      Height          =   372
      Left            =   1560
      TabIndex        =   21
      Top             =   480
      Width           =   1692
      _ExtentX        =   2990
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
   Begin VB.Frame pat_det 
      Caption         =   "PATIENT _ DETAILS"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   9132
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
         Height          =   285
         Left            =   6120
         TabIndex        =   66
         Top             =   945
         Visible         =   0   'False
         Width           =   1185
      End
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
         Height          =   285
         Left            =   7425
         TabIndex        =   65
         Top             =   945
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lbl_typ 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   7560
         TabIndex        =   30
         Top             =   600
         Width           =   1212
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   192
         Left            =   6804
         TabIndex        =   29
         Top             =   600
         Width           =   444
      End
      Begin VB.Label lbl_tr_no 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   4920
         TabIndex        =   28
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label lbl_cr_card_no 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   4920
         TabIndex        =   27
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label lbl_trust_no 
         AutoSize        =   -1  'True
         Caption         =   "Trust A/C no."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   192
         Left            =   3360
         TabIndex        =   26
         Top             =   960
         Width           =   1104
      End
      Begin VB.Label lbl_cr_no 
         AutoSize        =   -1  'True
         Caption         =   "Credit_card_no."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   192
         Left            =   3360
         TabIndex        =   24
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label lbl_cat 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   7560
         TabIndex        =   22
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label lbl_age 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   1440
         TabIndex        =   20
         Top             =   960
         Width           =   372
      End
      Begin VB.Label lbl_sex 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   252
         Left            =   1440
         TabIndex        =   19
         Top             =   600
         Width           =   372
      End
      Begin VB.Label lbl_name 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   252
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   4692
      End
      Begin VB.Label lbl_age1 
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
         ForeColor       =   &H00404080&
         Height          =   192
         Left            =   1044
         TabIndex        =   17
         Top             =   960
         Width           =   348
      End
      Begin VB.Label lbl_cat1 
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
         ForeColor       =   &H00404080&
         Height          =   192
         Left            =   6480
         TabIndex        =   16
         Top             =   240
         Width           =   768
      End
      Begin VB.Label lbl_sex1 
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
         ForeColor       =   &H00404080&
         Height          =   192
         Left            =   1068
         TabIndex        =   15
         Top             =   600
         Width           =   324
      End
      Begin VB.Label lbl_name1 
         AutoSize        =   -1  'True
         Caption         =   "Patient's Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   192
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1272
      End
      Begin VB.Label lbl_yrs 
         Caption         =   "Yrs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   252
         Left            =   1920
         TabIndex        =   12
         Top             =   960
         Width           =   372
      End
   End
   Begin MSMask.MaskEdBox me_rdate 
      Height          =   252
      Left            =   7560
      TabIndex        =   56
      Top             =   3240
      Width           =   1452
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
   Begin VB.ComboBox cbo_reqno 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3240
      Width           =   3495
   End
   Begin VB.TextBox txt_req_no 
      Height          =   285
      Left            =   1800
      MaxLength       =   13
      TabIndex        =   55
      Top             =   3240
      Width           =   3255
   End
   Begin VB.CommandButton New1 
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
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label lbl_csnr 
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
      Height          =   285
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label17 
      Caption         =   "NET PAYABLE AMOUNT:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4920
      TabIndex        =   64
      Top             =   5520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "PROTOCOL AMOUNT:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   360
      TabIndex        =   62
      Top             =   5520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "PROTOCOL PATIENT"
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
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6240
      TabIndex        =   60
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C000C0&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   -120
      X2              =   9360
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "&Entered By:"
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
      Height          =   195
      Left            =   6120
      TabIndex        =   59
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
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
      Height          =   192
      Left            =   156
      TabIndex        =   58
      Top             =   3240
      Width           =   1608
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
      Height          =   252
      Left            =   5640
      TabIndex        =   57
      Top             =   3240
      Width           =   1812
   End
   Begin VB.Label lbl_dno 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5880
      TabIndex        =   53
      Top             =   840
      Width           =   3372
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
      Height          =   372
      Left            =   4320
      TabIndex        =   52
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   192
      Left            =   240
      TabIndex        =   51
      Top             =   840
      Width           =   1248
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
      Height          =   252
      Left            =   3240
      TabIndex        =   50
      Top             =   2880
      Width           =   2052
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
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
      Height          =   192
      Left            =   4440
      TabIndex        =   48
      Top             =   840
      Width           =   1368
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   192
      Left            =   2880
      TabIndex        =   47
      Top             =   1320
      Width           =   1248
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL (Rs.):"
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
      Height          =   300
      Left            =   5520
      TabIndex        =   46
      Top             =   6000
      Width           =   1590
   End
   Begin VB.Label lbl_tot_amt 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lbl_dcode 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   1800
      TabIndex        =   8
      Top             =   2880
      Width           =   1092
   End
   Begin VB.Label lbl_dname 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   1212
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C000C0&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   -120
      X2              =   9360
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label lbl_doc_code 
      AutoSize        =   -1  'True
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
      Height          =   192
      Left            =   120
      TabIndex        =   45
      Top             =   2880
      Width           =   1644
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000B&
      Caption         =   "Remarks :"
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
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "OUT-PATIENT MEMO OF CHARGES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3600
      TabIndex        =   41
      Top             =   240
      Width           =   3375
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
      TabIndex        =   40
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   192
      Left            =   240
      TabIndex        =   39
      Top             =   600
      Width           =   1248
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
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
      Height          =   195
      Left            =   5520
      TabIndex        =   37
      Top             =   600
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C000C0&
      Height          =   7845
      Left            =   0
      Top             =   0
      Width           =   9495
   End
   Begin VB.Line Line2 
      X1              =   -120
      X2              =   9360
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1560
      TabIndex        =   44
      Top             =   0
      Width           =   7770
   End
End
Attribute VB_Name = "op_memo_d1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'31/03/2003-Parag/Sandhya
'Mapping Patch w.f.to New Schedule of Charges APR-2003.

Dim ZUSESERCODEMAP As Boolean
Dim ZUSESERINDMAP As Boolean
Dim RS33QR As rdoQuery
Dim zsercodemap As rdoQuery
Dim zrssercodemap As rdoResultset

Dim zserindmap As rdoQuery
Dim zrsserindmap As rdoResultset

Dim zstr As String

Dim patndt_qr As rdoQuery
Dim patndt_rs As rdoResultset

Dim zmemorcpt As rdoQuery
Dim zrsmemorcpt As rdoResultset
Dim zrsmemorcptflag As Boolean
Dim REQGOTFALG As Boolean
Dim RQ_SERIES As rdoQuery
Dim RS_SERIES As rdoResultset
Dim RS_SERIESflag As Boolean
Dim K As String                     ' ADDED 11/11/2001
Dim k1 As String                    ' ADDED 11/11/2001

Dim SEL_DISICU_PM_QR As rdoQuery 'Shraddha 27/10/12
Dim SEL_DISICU_PM_RS As rdoResultset 'Shraddha 27/10/12

Dim S
Dim rq_updtSERIES As rdoQuery
Dim rq_reqdet As rdoQuery
Dim RS_reqdet As rdoResultset
Dim RS_reqdetflag As Boolean
Dim rq_updtreqdet As rdoQuery
Dim rq_updtdelreqdet As rdoQuery
Dim RQ_REQDETREQNO As rdoQuery
Dim RS_REQDETREQNO As rdoResultset
Dim RS_REQDETREQNOflag As Boolean
Dim REQUISITN_UPDT As rdoQuery

Dim disreq_updt As rdoQuery

Dim RQ_reqdet1 As rdoQuery
Dim rs_reqdet1 As rdoResultset
Dim RS_reqdet1flag As Boolean
Dim UPDTA107_QR As rdoQuery

Dim requisitn_qr As rdoQuery
Dim requisitn_rs As rdoResultset
Dim requisitn_rsflag As Boolean
'*******
Dim zpattype As String
Dim CASENO As String
Dim flag1 As Integer
Dim FLEX2FLAG As Boolean
Dim CANCELFLG  '***31399r
Dim NoCaseFlag '***3499r
Dim EmptyGrid '***3499r
Dim AmtNotEq As Integer  ''''25499r
Dim iRow As Integer
Dim DOCNOTFOUND As Boolean
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
Dim rs22 As rdoResultset
Dim rs22flag As Boolean
Dim rs33 As rdoResultset
Dim rs33flag As Boolean
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
'Dim RS17 As rdoResultset
Public RS17 As rdoResultset
Dim rs17flag As Boolean
'RS17 As rdoResultset
Dim RS18 As rdoResultset
Dim rs18flag As Boolean
Dim rs2 As rdoResultset
Dim rs2flag As Boolean
Dim rs_dname As rdoResultset
Dim rs_dnameflag As Boolean
Dim a As Integer
Dim ncchg As String  ' added for z-series memos***4899

Dim RS As rdoResultset
Dim rsflag As Boolean
Dim flag As Integer
Private mTotalRows& ' Contains the total rows in the set of records
Private UserData() As Variant ' 2-dimensional array containing records
Private Const MAXCOLS = 6 ' Maximum number of fields in record set.
Dim reflg As Integer
Dim FLGTA As Integer '**17/12/98
Dim FLGCR As Integer '**17/12/98

Dim Ta107_InsrtQr As rdoQuery
Dim Ta1071_InsrtQr As rdoQuery
Dim Ta113_UpdtQr As rdoQuery
Dim Ta114_UpdtQr As rdoQuery
Dim Scrn_AuthQr As rdoQuery
Dim Doc_ViewQr As rdoQuery
Dim Ma1031Qr As rdoQuery
Dim Ma106Qr As rdoQuery
Dim MA102DOC As rdoQuery
Dim Ma102Qr As rdoQuery
Dim Ta113Qr As rdoQuery
Dim DELTA107 As rdoQuery
Dim DELTA1071 As rdoQuery

 Dim Ta114Qr As rdoQuery
Dim Ta107Qr As rdoQuery
Dim Ta1071Qr As rdoQuery
Dim SELTA1071_QR As rdoQuery
Dim SELTA1071_RS As rdoResultset
Dim selta1071_rsflag As Boolean
Dim AllMa101_Qr As rdoQuery
Dim SELRREQQR As rdoQuery
Dim SELRREQRS As rdoResultset
Dim SELRREQRSFLAG As Boolean
Dim InPt_ChkQr As rdoQuery
Dim Exist_ChkQr As rdoQuery
Dim Ma101Qr As rdoQuery
Dim OPQ As rdoQuery '''A251199

Dim INNC As Boolean

Dim PrintResult%

Dim SELDMULPROCEQR As rdoQuery
Dim SELDMULPROCERS As rdoResultset
Dim SELDMULPROCEFLAG As Boolean

Dim SELDMOLEPROCQR As rdoQuery
Dim SELDMOLEPROCRS As rdoResultset
Dim SELDMOLEPROCFLAG As Boolean

Dim zgetservice As rdoQuery
Dim zrsgetservice As rdoResultset
Dim zrsgetserviceflag As Boolean
'***** MAPING THE SERVICE CODE WITH DOCTOR CODE
Dim zmaptable As rdoQuery
Dim zrsmaptable As rdoResultset
Dim ZRSMAPTABLEFLAG As Boolean

Dim zupdtmap As rdoQuery

Dim zta1071 As rdoQuery
Dim zrsta1071 As rdoResultset
Dim Zrsta1071flag As Boolean
'***** MAPING THE SERVICE CODE WITH DOCTOR CODE

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

'V001 PATIENT DEPOSIT STARTS

    Dim PATDEPO  As Boolean
    Dim TA116Qr As rdoQuery
    Dim Ta116_UpdtQr As rdoQuery
'V001 PATIENT DEPOSIT STARTS

'26/10/2005--BIJESH
Dim ZAS400DTTM As rdoResultset

''' SANDHYA DATED 12/3/2013
Dim zgetSerA001 As rdoQuery
Dim zrsgetSerA001 As rdoResultset

Dim zgetSerCodes As rdoQuery
Dim zrsgetSercodes As rdoResultset
Dim memoduplicate As Boolean
''' 1.8.2008 - SANDHYA protocol patient starts
Dim PROTFLG As Boolean
Dim NPROTFLG As Boolean
Dim PROTCASE_QR As rdoQuery
Dim PROTCASE_RS As rdoResultset
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
Dim PINAME As String
Dim RSPINAME As rdoResultset

'***** Shraddha 27/07/10 *****
Dim SELDISWARD_QR As rdoQuery
Dim SELDISWARD_RS As rdoResultset
'''' added by shraddha
'''' dated 16.8.2010 to check if service code exist in table ser_ref for duplication
Dim qrser_ref As rdoQuery
Dim rsser_ref As rdoResultset

Dim scard As Boolean '''added by charulata
Dim scard_bal_amt As Double  '''added by charulata for smartcard
Dim PCARD As PISWriteCard.ClsWriteCard  '''added by charulata
Dim PINFO As PatientInformation  '''added by charulata
'''' ADDED BY SANDHYA DATED 11.5.2011
'''' access docname from empdetail table
Dim SELDRCODE As rdoResultset
Dim rsdocview As rdoResultset
Dim SELMA102QR As rdoQuery
Dim SELMA102RS As rdoResultset
Dim SELDOCCODE As rdoQuery
Dim RSSELDCODE As rdoResultset
Dim MDRCODE As String
'''' added sandhya dated FOR ACTREC 11.6.2012
Dim actmemflg As Boolean
''''' ADDED BY SANDHYA DATED 6.07.2012 FOR BP PATIENT RAJIV GANDHI

'Dim bppatflag As Boolean
Dim QRBPPATN As rdoQuery
Dim RSBPPATN As rdoResultset
Dim UPDbppatn As rdoQuery
Dim INS_BPPATMEMO As rdoQuery


'*****************************

'' ADDED BY SANDHYA PROTOCOL PATIENT
Public Sub NPRTMEMOCR()
Dim STR5 As String
Dim STR6 As String
Dim STR7 As String
Dim STR8 As String
Dim STR9 As String
Dim STR10 As String
Dim STR11 As String
Dim str12 As String
'Dim CHKAMT As Integer
Dim CHKAMT As Double
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
'''ipcr.ReportFileName = "C:\PB121298\OPMEMO.RPT"
'''ipcr.Action = 1

ipcr.ReportFileName = "C:\PB121298\ONPRTMEMO.rpt"
ipcr.SelectionFormula = "{protmemo.MEMO_NO} = '" & me_mno.Text & "'"
' For J = 0 To 130 '261099
'    ipcr.Formulas(J) = " "
'Next J

If flag = 1 Then
    ipcr.Formulas(0) = "fdeptname = '" & Com_dno.Text & "'"
    'ipcr.Formulas(2) = "fcaseno = '" & CASENO & "'"
    Exist_ChkQr(0) = me_mno.Text    '"M-" & me_mno.Text
    str12 = CASENO
    Set rs7 = Exist_ChkQr.OpenResultset(1, 1)
    STR10 = "*A*"
Else
    ipcr.Formulas(0) = "fdeptname = '" & lbl_dno & "'"
    Exist_ChkQr(0) = M_number '"M-" & M_number
    str12 = lbl_cno
    Set rs7 = Exist_ChkQr.OpenResultset(1, 1)
    STR10 = "*V*"
End If

'If Label12.Caption = "DOCTOR NAME : " Then
'     ipcr.Formulas(1) = "doc_name = " & a & txt_unit_doc_name & a & ""
'     ipcr.Formulas(2) = "doc_code = " & a & lbl_dcode & a & ""
' ElseIf Label12.Caption = "UNIT NAME : " Then
'     ipcr.Formulas(1) = "doc_name = " & a & txt_unit_doc_name & a & ""
'     ipcr.Formulas(2) = "doc_code = " & a & 0 & a & ""
' Else
     ipcr.Formulas(1) = "doc_name = " & a & txt_unit_doc_name & a & ""
     ipcr.Formulas(2) = "doc_code = " & a & 0 & a & ""
'End If
''IPCR.Formulas(2) = "Fcaseno = '" & CASENO & "'"
'
''ipcr.Formulas(3) = "fdcode = " & a & lbl_dname & a & ""
ipcr.Formulas(4) = "NAME_CON = " & a & Trim(UCase(LBL_NAME)) & a & ""
ipcr.Formulas(5) = "ftot = " & a & " *****" & Format(lbl_tot_amt, "#,##,##,##,##0.00") & a & ""
ipcr.Formulas(6) = "Ptot = " & a & " *****" & Format(prottot, "#,##,##,##,##0.00") & a & ""
ipcr.Formulas(7) = "nettot = " & a & " *****" & Format(nettot, "#,##,##,##,##0.00") & a & ""
'
'
'STR1 = RTrim(username) 'change this var
 STR5 = RTrim(usr_name)
 STR6 = Format(DATE11, "DD/MM/YYYY")
 STR7 = time1
 'Exist_ChkQr(0) = "M-100010"
'Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
If rs7.RowCount <> 0 Then
    STR9 = Trim(rs7("D_S_NAME"))
End If
 'STR8 = Trim(STR5) + "  " + Trim(STR6) + "  " + Trim(STR7)
 STR8 = Trim(STR9) + " / " + Trim(STR5) + " " + Trim(STR10) + " " + " " + STR11 + " " + Trim(STR6) + "  " + Trim(STR7)
ipcr.Formulas(8) = "us_dt_tm = " & a & STR8 & a & ""
'
 If rs2("PAT_TYPE") = "TA" Then
     ipcr.Formulas(9) = "cr_tr = '" & rs_tr("TR_ACC_NO") & "'"
     ipcr.Formulas(10) = "PATTP = '" & "TRUST PATIENT" & "'"
 ElseIf rs2("PAT_TYPE") = "CR" Then
    ipcr.Formulas(9) = "cr_tr = '" & Trim(rs_cr("CR_CARD")) & "'"
 ElseIf rs2("PAT_TYPE") = "CP" And PATDEPO = True Then
    ipcr.Formulas(9) = "cr_tr = '" & rs_tr("PD_ACC_NO") & "'"
    ipcr.Formulas(10) = "PATTP = '" & "DEPOSIT PATIENT" & "'"
 ElseIf PTOT > 0 Then
    ipcr.Formulas(9) = "cr_tr = '" & "    " & "'"
    ipcr.Formulas(10) = "PATTP = '" & "PROTOCOL PATIENT" & "'"
 Else
    ipcr.Formulas(9) = "cr_tr = '" & "    " & "'"
    ipcr.Formulas(10) = "PATTP = '" & " " & "'"
 End If



    'DBGrid1.Refresh




ipcr.CopiesToPrinter = 1
  
''Dim P As Object
For Each p In Printers


'IPCR.PrinterName = "Wipro LQ-1050+DX"
'ipcr.WindowState = crptMaximized
ipcr.Destination = crptPRINTER
'PrintResult% = ipcr.PrintReport
ipcr.Action = 1
If PrintResult% <> 0 Then
MsgBox Error(Err)
End If
Exit For
Next

End Sub

''*****FOR BP CROSSMATCH******


'**1/7/99
Sub LVL_BTN_ENABL() '****12/6/99

Dim RS_LVL As rdoResultset
Dim RS_FM As rdoResultset

'Set RS_LVL = rdb.OpenResultset("select * from TASEC2 where TRIM(USR_GRP) = '" & Trim(USRGRP) & "' AND TRIM(SCREEN) = '" & UCase(Trim(reg_doc.Name)) & "'", 2, 3)
Scrn_AuthQr(0) = Trim(USRGRP)
Scrn_AuthQr(1) = UCase(Trim(Me.name))
Set RS_LVL = Scrn_AuthQr.OpenResultset(2, 3)
''''Set RS_LVL = rdb.OpenResultset("select * from TOOLSADMN.scrn_auth1 where TRIM(USR_GRP) = '" & Trim(USRGRP) & "' AND TRIM(SCRN_ID) = '" & UCase(Trim(Me.Name)) & "'", 2, 3)
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
    Update.Enabled = True
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
            'Update.Visible = True
            View.Visible = False
            'delete.Visible = True '9/6/99 COMMENTED BECAUSE NO DELETE OPTION IS PROVIDED
            exit1.Visible = False
            Back.Visible = True
            commit.Visible = False
            cancel.Visible = False
            print1.Visible = True
        End If
        
End Sub

'Sub del()
'Dim i
'Dim MemNo As String
'If flag1 = 3 Then
'    i = MsgBox("Do you want to delete ? ", vbYesNo)
'    If i = vbYes Then
'        Set rs30 = rdb.OpenResultset("select * from ta107 where memo_no = '" & "M-" & M_number.Text & "'", rdOpenKeyset, 3)
'        Set rs31 = rdb.OpenResultset("select * from TA1071 where memo_no = '" & "M-" & M_number.Text & "'", rdOpenKeyset, 3)
'        '***********************020399
'        ' MsgBox rs30("AMT_RCVD")
'        Select Case rs30("AMT_RCVD")
'        Case "C"
''            If ACCOUNTPSWD = "ACCOUNTS" Then
''                MsgBox "PATIENT IS COMPANY REFFERED."
''            Else
'                MsgBox "PATIENT IS COMPANY REFFERED, YOU CAN'T DELETE"
'                Call Refresh1
'                delete.Visible = False
'                Exit Sub
''            End If
'        Case "T"
''            If ACCOUNTPSWD = "ACCOUNTS" Then
''                MsgBox "PATIENT IS TRUST AIDED."
''            Else
'                MsgBox "PATIENT IS TRUST AIDED, YOU CAN'T DELETE"
'                Call Refresh1
'                delete.Visible = False
'                Exit Sub
''            End If
'        Case "I"
'            MsgBox "INTERIM BILL IS GENERATED FOR THIS MEMO, YOU CAN'T DELETE"
'            Call Refresh1
'            delete.Visible = False
'            Exit Sub
'        Case "F"
'            MsgBox "FINAL BILL IS GENERATED FOR THIS MEMO, YOU CAN'T DELETE"
'            Call Refresh1
'            delete.Visible = False
'            Exit Sub
'         Case "R"
'            MsgBox "RECEIPT IS GENERATED FOR THIS MEMO, YOU CAN'T DELETE"
'            Call Refresh1
'            delete.Visible = False
'            Exit Sub
'        End Select
'MemNo = "M-" & M_number.Text
''        rs30.delete
''NEW DELETE CODE WRITTEN
'
'rdb.BeginTrans
'On Error GoTo DELERROR
'DELTA107(0) = MemNo
'DELTA107.Execute
'
'DELTA1071(0) = MemNo
'DELTA1071.Execute
'
'rdb.CommitTrans
'
'
' MsgBox "RECORD HAS BEEN DELETED ..."
' delete.Visible = False
' cleartxt
'        M_number = " "
'        lbl_cno = " "
'        ReDim UserData(0 To 5, 0 To 12)
'        DBGrid1.Refresh
'        New1.SetFocus
' Exit Sub
'
'
''
''         While Not rs31.EOF
''            rs31.delete
''         If Not rs31.EOF Then
''            rs31.MoveNext
''         End If
''         Wend
'
'        '**********************7499
'
'        '**********************7499
'    Else
'DELERROR:
'
'        MsgBox "RECORD NOT DELETED ..."
'        delete.Visible = False
'    End If
'End If
'End Sub
Sub ud_name()
d_u_name.clear
If lbl_cat = "C " Or lbl_cat = "NC" Then
'''    Set RS17 = rdb.OpenResultset("select ud_name from MA1031 where u_d = 'U'", rdOpenKeyset)
    RS17.MoveFirst
    While Not RS17.EOF
            If RS17("U_D") = "U" Then
                 d_u_name.AddItem RS17("ud_name")
            End If
            RS17.MoveNext
    Wend
Else
''' ADDED BY SANDHYA DATED 11.5.2011
''''' LISTING DOCTORS FROM PAYROLLLIB WHOSE CESSYN = 4
  SELMA102RS.Requery
  If SELMA102RS.RowCount > 0 Then
    While Not SELMA102RS.EOF
         d_u_name.AddItem Trim(SELMA102RS("DOC_NAME"))
         SELMA102RS.MoveNext
    Wend
  End If
'    RS17.MoveFirst
'    While Not RS17.EOF
'            If RS17("U_D") = "D" Then
'                 d_u_name.AddItem RS17("ud_name")
'            End If
'            RS17.MoveNext
'    Wend

'''    Set RS17 = rdb.OpenResultset("select ud_name from MA1031 where u_d = 'D'", rdOpenKeyset)
End If
End Sub
Sub dcode()
Dim opt As String

opt = lbl_cat
    
    Select Case opt
    
    Case "C ", "NC"
      'MsgBox " Referred to Unit "
      If rs2("pat_reg_ty") = "R" Then '---121099
          Call doc_code
      End If '---121099
    Case "A ", "B ", "D ", "F "
      
      Call doc_code
      
    End Select
    
End Sub


Sub pat_type()
If rs2("pat_reg_ty") = "G" Then
   Label12.Caption = "UNIT NAME : "
   d_u_name.Visible = False
   txt_unit_doc_name.TabStop = False
   If flag = 1 Then
       DBGrid1.SetFocus
   End If
   lbl_dcode = "  "
ElseIf rs2("pat_reg_ty") = "P" Then
    Label12.Caption = "DOCTOR NAME : "
    If flag = 1 Then
        d_u_name.Visible = True
        Call doc_code
    End If
    
'    lbl_dcode = rs33("doc_code")
    'lbl_dcode = 11016 'rs33("doc_code")
End If

If rs2("pat_reg_ty") = "R" Then
 If rs2("U_D") = "D" Then
    Label12.Caption = "DOCTOR NAME : "
    If flag = 1 Then
        d_u_name.Visible = True
    End If
    If flag = 1 Then
    Call doc_code
    End If
 ElseIf rs2("U_D") = "U" Then
       Label12.Caption = "UNIT NAME : "
       d_u_name.Visible = False
       txt_unit_doc_name.TabStop = False
       If flag = 1 Then
         DBGrid1.SetFocus
       End If
   lbl_dcode = "  "
    
'    lbl_dcode = rs33("doc_code")
    'lbl_dcode = 11016 'rs33("doc_code")
End If
End If
End Sub
Sub patient_det(case_num As String)
Dim d As Date
Dim D1 As Date
On Error GoTo MSG:
pat_det.Visible = True

'''Set rs2 = rdb.OpenResultset("select first_name,mid_name,last_name,birth_date,sex,pat_catg,pat_reg_ty,pat_type,ud_name,consult_dt,u_d from TA101 where case_no = '" & case_num & "'", rdOpenKeyset)
If flag = 2 Then '''A251199
    If rs2flag = False Then
        OPQ(0) = case_num
        Set rs2 = OPQ.OpenResultset(1, 3)
        rs2flag = True
    Else
        OPQ(0) = case_num
        rs2.MoreResults
        rs2.Requery
    End If
    
    
    
'Set Ta107Qr = rdb.CreateQuery("Ta107Q", STR)
If rs10flag = False Then
    Ta107Qr(0) = M_number.Text ' "M-" & M_number
    Set rs10 = Ta107Qr.OpenResultset(1, 1)
    rs10flag = True
Else
    Ta107Qr(0) = M_number.Text
    rs10.MoreResults
    rs10.Requery
End If
'Set MA102DOC = rdb.CreateQuery("Ma102DOCQ", "select DOC_NAME from MA102 where doc_code = ? ")
 If rs_dnameflag = False Then
     MA102DOC(0) = rs10("DOC_CODE")
    Set rs_dname = MA102DOC.OpenResultset(1, 3)
    rs_dnameflag = True
 Else
    MA102DOC(0) = rs10("DOC_CODE")
    rs_dname.MoreResults
    rs_dname.Requery
 End If
If rs_dname.RowCount = 0 Then
txt_unit_doc_name = rs2("ud_name")
'txt_unit_doc_name = "0"
Else
txt_unit_doc_name = Trim(rs_dname("doc_name"))
End If

Else
txt_unit_doc_name = Trim(rs2("UD_NAME"))
  
     ''''Set rs2 = OPQ.OpenResultset(rdOpenKeyset)
End If
'''Set rs_cr = rdb.OpenResultset("select COUNT(*) as c from ta114 where case_no = '" & case_num & "'", rdOpenKeyset)
'''Set rs_tr = rdb.OpenResultset("select COUNT(*) as c from ta113 where case_no = '" & case_num & "'", rdOpenKeyset)
'txt_unit_doc_name = rs2("ud_name")
Call pat_type

LBL_NAME = RTrim(rs2("first_name")) & " " & RTrim(rs2("mid_name")) & " " & rs2("last_name")
lbl_sex = rs2("sex")
'txt_unit_doc_name = "S.D.BANAVALI"
If flag = 1 Then   '9/6/99 ADDED BECAUSE TO VIEW IP/OP MEMOS PATIENT CATG SHOULD BE THE CATEGORY AT THE TIME OF MEMO RAISING WHERE AS TA101 CATG IS LATEST CATG
   lbl_cat = rs2("pat_catg") '9/6/99 ADDED BECAUSE TO VIEW IP/OP MEMOS PATIENT CATG SHOULD BE THE CATEGORY AT THE TIME OF MEMO RAISING WHERE AS TA101 CATG IS LATEST CATG
End If '9/6/99 ADDED BECAUSE TO VIEW IP/OP MEMOS PATIENT CATG SHOULD BE THE CATEGORY AT THE TIME OF MEMO RAISING WHERE AS TA101 CATG IS LATEST CATG
'lbl_cat = rs2("pat_catg") '9/6/99 COMMENTED BECAUSE TO VIEW IP/OP MEMOS PATIENT CATG SHOULD BE THE CATEGORY AT THE TIME OF MEMO RAISING WHERE AS TA101 CATG IS LATEST CATG
lbl_typ = rs2("pat_type") '***22/10/98

If rs2("PAT_TYPE") = "CR" Then
    If rs_crflag = False Then
        Ta114Qr(0) = case_num
        Set rs_cr = Ta114Qr.OpenResultset(2, 3)
        rs_crflag = True
    Else
        Ta114Qr(0) = case_num
        rs_cr.MoreResults
        rs_cr.Requery
    End If
    ''''Set rs_cr = rdb.OpenResultset("select * from ta114 where case_no = '" & case_num & "'", rdOpenKeyset)
    '''If rs_cr("c") = 1 Then
    If rs_cr.RowCount > 0 Then
        '''Set rs_cr = rdb.OpenResultset("select * from ta114 where case_no = '" & case_num & "'", rdOpenKeyset)
        lbl_cr_card_no = rs_cr("cr_card")
    Else
        MsgBox "Corresponding Credit Card No. does not exist...."
    End If
ElseIf rs2("PAT_TYPE") = "TA" Then
'       Ta114Qr(0) = case_num
        If rs_trflag = False Then
            Ta113Qr(0) = case_num
            Set rs_tr = Ta113Qr.OpenResultset(1, 3)
            rs_trflag = True
        Else
            Ta113Qr(0) = case_num
            rs_tr.MoreResults
            rs_tr.Requery
        End If
       ''''Set rs_tr = rdb.OpenResultset("select * from ta113 where case_no = '" & case_num & "'", rdOpenKeyset)
        '''If rs_tr("c") = 1 Then
        If rs_tr.RowCount > 0 Then
            '''Set rs_tr = rdb.OpenResultset("select * from ta113 where case_no = '" & case_num & "'", rdOpenKeyset)
            lbl_tr_no = rs_tr("tr_acc_no")
        Else
            MsgBox "Corresponding Trust Account No. does not exist...."
        End If
Else

   'V001 PATIENT DEPOSIT STARTS
    TA116Qr(0) = case_num
    Set rs_tr = TA116Qr.OpenResultset(2, 3)
    If rs_tr.RowCount = 1 Then
        PATDEPO = True
        lbl_cr_card_no = rs_tr("PD_ACC_NO")
        lbl_tr_no = " "
    Else
        lbl_cr_card_no = " " ' OLD
        lbl_tr_no = " " 'OLD
    End If
   'V001 PATIENT DEPOSIT ENDS
   
End If

'' SANDHYA 1.8.2008 PROTOCOL PATIENT
PROTCASE_QR(0) = case_num
PROTCASE_QR(1) = MIRBNO
Set PROTCASE_RS = PROTCASE_QR.OpenResultset(2, 3)
If PROTCASE_RS.RowCount > 0 Then
   PROTFLG = True
  ' lbl_tr_no = PROTCASE_RS("PCODE")
   Label15.Enabled = True
   Label15.Visible = True
Else
   PROTFLG = False
   Label15.Enabled = False
   Label15.Visible = False
End If

d = rs2("birth_date")
D1 = Format(d, "dd/mm/yyyy")
lbl_age = Abs(Right(DateDiff("yyyy", me_date, D1), 2)) '5599
'lbl_age = Right(DateDiff("yyyy", d1, me_date), 2) '5599
Exit Sub
MSG:  MsgBox Error(Err)
'me_cno.SetFocus
'me_cno.SelStart = 0
End Sub
Sub opmemocr()
Dim STR5 As String
Dim STR6 As String
Dim STR7 As String
Dim STR8 As String
Dim STR9 As String
Dim STR10 As String
Dim STR11 As String
Dim str12 As String
'Dim CHKAMT As Integer
Dim CHKAMT As Double
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
'''ipcr.ReportFileName = "C:\PB121298\OPMEMO.RPT"
'''ipcr.Action = 1
ipcr.ReportFileName = "C:\PB121298\OPMEMO.rpt"

For j = 0 To 130 '261099
    ipcr.Formulas(j) = " "
Next j

If flag = 1 Then
    If rs7flag = False Then
        ipcr.Formulas(0) = "fmemono = '" & me_mno.Text & "'"  '"M-" & me_mno.Text & "'"
        ipcr.Formulas(3) = "fdeptname = '" & Com_dno.Text & "'"
        ipcr.Formulas(2) = "fcaseno = '" & CASENO & "'"
        Exist_ChkQr(0) = me_mno.Text    '"M-" & me_mno.Text
        str12 = CASENO
        Set rs7 = Exist_ChkQr.OpenResultset(1, 1)
        STR10 = "*A*"
        rs7flag = True
    Else
        ipcr.Formulas(0) = "fmemono = '" & me_mno.Text & "'"  '"M-" & me_mno.Text & "'"
        ipcr.Formulas(3) = "fdeptname = '" & Com_dno.Text & "'"
        ipcr.Formulas(2) = "fcaseno = '" & CASENO & "'"
        Exist_ChkQr(0) = me_mno.Text    '"M-" & me_mno.Text
        str12 = CASENO
        rs7.MoreResults
        rs7.Requery
        STR10 = "*A*"
    End If
Else
    If rs7flag = False Then
        ipcr.Formulas(0) = "fmemono = '" & M_number.Text & "'"     '"M-" & M_number & "'"
        ipcr.Formulas(2) = "fcaseno = '" & lbl_cno & "'"
        ipcr.Formulas(3) = "fdeptname = '" & lbl_dno & "'"
        Exist_ChkQr(0) = M_number '"M-" & M_number
        str12 = lbl_cno
        Set rs7 = Exist_ChkQr.OpenResultset(1, 1)
        STR10 = "*V*"
        rs7flag = True
    Else
        ipcr.Formulas(0) = "fmemono = '" & M_number.Text & "'"     '"M-" & M_number & "'"
        ipcr.Formulas(2) = "fcaseno = '" & lbl_cno & "'"
        ipcr.Formulas(3) = "fdeptname = '" & lbl_dno & "'"
        Exist_ChkQr(0) = M_number '"M-" & M_number
        str12 = lbl_cno
        rs7.MoreResults
        rs7.Requery
        STR10 = "*V*"
    End If
End If

If Label12.Caption = "DOCTOR/UNIT NAME :" Then
     ipcr.Formulas(1) = "doc_name = " & a & Trim(txt_unit_doc_name) & a & ""
     ipcr.Formulas(8) = "doc_code = " & a & (lbl_dcode.Caption) & a & ""
 ElseIf Label12.Caption = "UNIT NAME : " Then
     ipcr.Formulas(1) = "doc_name = " & a & Trim(txt_unit_doc_name) & a & ""
     ipcr.Formulas(8) = "doc_code = " & a & 0 & a & ""
End If

'IPCR.Formulas(2) = "Fcaseno = '" & CASENO & "'"

ipcr.Formulas(4) = "fdcode = " & a & lbl_dname & a & ""
ipcr.Formulas(5) = "NAME_CON = " & a & Trim(UCase(LBL_NAME)) & a & ""
ipcr.Formulas(6) = "fpat_catg = '" & lbl_cat & "'" 'modify

'''
'IPCR.Formulas(7) = "fward = '" & lbl_wf_bed & "'" 'modify
'IPCR.Formulas(9) = "freqno = " & A & Trim(txt_req_no.Text) & A & ""
'IPCR.Formulas(11) = "freqdate = " & A & Trim(me_rdate.Text) & A & ""
'''

'IPCR.Formulas(8) = "acc_num = " & A & "lbl_tr_no" & A & ""
'IPCR.Formulas(9) = "acc_head = " & A & Trim(lbl_acc_head) & A & ""

ipcr.Formulas(10) = "ftot = " & a & " *****" & Format(lbl_tot_amt, "#,##,##,##,##0.00") & a & ""

ipcr.Formulas(12) = "fremark = " & a & Trim(txt_rem.Text) & a & ""
ipcr.Formulas(13) = "fmemodate  = " & a & Trim(me_date.Text) & a & ""
'STR1 = RTrim(username) 'change this var
 STR5 = RTrim(usr_name)
 STR6 = Format(DATE11, "DD/MM/YYYY")
 STR7 = time1
 'Exist_ChkQr(0) = "M-100010"
'Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
If rs7flag = False Then
        Ta107Qr(0) = me_mno.Text   ' "M-" & me_mno.Text     'BILLOPT 61299 ADDED
        Ta107Qr(1) = Trim(pathology.txt_status.Text)
        Set rs7 = Ta107Qr.OpenResultset(1, 3)
        'rs7flag = True
    Else
        Ta107Qr(0) = me_mno.Text   ' "M-" & me_mno.Text     'BILLOPT 61299 ADDED
        Ta107Qr(1) = Trim(pathology.txt_status.Text)
        Set rs7 = Ta107Qr.OpenResultset(1, 3)
End If

If rs7.RowCount <> 0 Then
    STR9 = Trim(rs7("D_S_NAME"))
End If
 'STR8 = Trim(STR5) + "  " + Trim(STR6) + "  " + Trim(STR7)
 STR8 = Trim(STR9) + " / " + Trim(STR5) + " " + Trim(STR10) + " " + " " + STR11 + " " + Trim(STR6) + "  " + Trim(STR7)
ipcr.Formulas(14) = "us_dt_tm = " & a & STR8 & a & ""
 '''' ADDED BY SANDHYA DATED 7/7/2012 FOR RGJAY PATIENT
If BPPATFLAG = True Then
  If flag = 1 Then
    Set RSBPPATN = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.BPPATIENT WHERE CASENO ='" & CASENO & "'", 1, 1)
  End If
  If flag = 2 Then
    Set RSBPPATN = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.BPPATIENT WHERE CASENO ='" & lbl_cno & "'", 1, 1)
  End If
End If

 If rs2("PAT_TYPE") = "TA" Then
     ipcr.Formulas(15) = "cr_tr = '" & rs_tr("TR_ACC_NO") & "'"
     ipcr.Formulas(150) = "PATTP = '" & "TRUST PATIENT" & "'"
  '''' ADDED BY SANDHYA DATED 7/7/2012 FOR RGJAY PATIENT
 ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = False Then
    ipcr.Formulas(15) = "cr_tr = '" & Trim(rs_cr("CR_CARD")) & "'"
    ipcr.Formulas(150) = "PATTP = '" & "COMPANY PATIENT" & "'"
 ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = True Then
     ipcr.Formulas(15) = "cr_tr = '" & Trim(RSBPPATN("BPCARD_NO")) & "'"
     ipcr.Formulas(150) = "PATTP = '" & "RGJAY PATIENT" & "'"
 ElseIf rs2("PAT_TYPE") = "CP" And PATDEPO = True Then
     ipcr.Formulas(15) = "cr_tr = '" & rs_tr("PD_ACC_NO") & "'"
     ipcr.Formulas(150) = "PATTP = '" & "DEPOSIT PATIENT" & "'"
ElseIf rs2("PAT_TYPE") = "CP" And scard = True Then
     ipcr.Formulas(15) = "cr_tr = '" & Right(me_cno.Text, 5) & Left(me_cno.Text, 2) & "'"
     ipcr.Formulas(150) = "PATTP = '" & "SMART CARD PATIENT" & "'"
 Else
     ipcr.Formulas(15) = "cr_tr = '" & "    " & "'"
     ipcr.Formulas(150) = "PATTP = '" & " " & "'"
 End If

i = 0
K = 0

If flag = 1 Then
   If rs7flag = False Then
        Ta1071Qr(0) = me_mno.Text   ' "M-" & me_mno.Text     'BILLOPT 61299 ADDED
        Set rs7 = Ta1071Qr.OpenResultset(1, 3)
        'rs7flag = True
    Else
        Ta1071Qr(0) = me_mno.Text   ' "M-" & me_mno.Text     'BILLOPT 61299 ADDED
        Set rs7 = Ta1071Qr.OpenResultset(1, 3)
    End If
Else
    If rs7flag = False Then
        Ta1071Qr(0) = M_number  '"M-" & M_number
         Set rs7 = Ta1071Qr.OpenResultset(1, 3)
    Else
        Ta1071Qr(0) = M_number
         Set rs7 = Ta1071Qr.OpenResultset(1, 3)
    End If
End If



DBGrid1.Refresh
rrow1 = 0
CHKAMT = 0
If rs7.RowCount Then
   rs7.MoveFirst
   For rrow1 = 0 To rs7.RowCount - 1
     ipcr.Formulas(17 + rrow1 + K) = "s" & rrow1 & " = " & a & Trim(rs7(1).Value) & a & ""
     If Not IsNull(rs7(1).Value) Then
        If rs13flag = False Then
            Ma101Qr(0) = rs7(1).Value
            Set rs13 = Ma101Qr.OpenResultset(2, 3)
            rs13flag = True
        Else
            Ma101Qr(0) = rs7(1).Value
            rs13.MoreResults
            rs13.Requery
        End If
        ipcr.Formulas(40 + rrow1 + K) = "ser_det" & rrow1 & " = " & a & Trim(rs13("ser_det")) & a & ""  'modify
        ipcr.Formulas(64 + rrow1 + K) = "q" & rrow1 & " = " & a & Trim(rs7("QTY")) & a & ""   'modify
        samt = Format((rs7(2).Value), "#,##,##,##,###0.00")
        samt = RTrim(samt)
        l = Len(samt)
        samt = Space(18 - l) + samt
        'CHKAMT = CHKAMT + samt
        CHKAMT = CHKAMT + CDbl(rs7(2))
        ipcr.Formulas(88 + rrow1 + K) = "ser_amt" & rrow1 & " = " & a & samt & a & "" 'modify
        If Trim(rs7("REQ_NO")) = "" Then
        ipcr.Formulas(112 + rrow1 + K) = "REQ" & rrow1 & " = " & a & "" & a & "" 'modify
        Else
        ipcr.Formulas(112 + rrow1 + K) = "REQ" & rrow1 & " = " & a & rs7("REQ_NO") & a & "" 'modify
        End If
     End If
        K = K + 1
        rs7.MoveNext
    Next rrow1
End If
    'DBGrid1.Refresh
If rs2flag = False Then
   OPQ(0) = str12
   Set rs2 = OPQ.OpenResultset(1, 3)
   rs2flag = True
Else
    OPQ(0) = str12
    rs2.MoreResults
    rs2.Requery
End If
'MsgBox rs2.RowCount
If rs2.RowCount > 0 Then
    If requisitn_rsflag = False Then
        requisitn_qr(0) = str12
        Set requisitn_rs = requisitn_qr.OpenResultset(1, 3)
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
'''' added by sandhya for actrec patient treated at tmh
'''' dated 11.6.2012
If actmemflg = True Then
   ipcr.Formulas(151) = "acttmhfl = '" & "Actrec Patient treated at TMH" & "'"
                                 
Else
   ipcr.Formulas(151) = "acttmhfl = '" & "         " & "'"
End If



If CDbl(lbl_tot_amt) = CHKAMT Then
    If rs2flag = False Then
        OPQ(0) = str12
        Set rs2 = OPQ.OpenResultset(1, 3)
        rs2flag = True
    Else
        OPQ(0) = str12
        rs2.MoreResults
        rs2.Requery
    End If
    'Shraddha 07/01/09 2 copies removed for credit patient
     '  ipcr.CopiesToPrinter = 1

'IPCR.WindowState = crptMaximized
'Report1.PrinterName = "Epson LQ850"
'IPCR.PrinterName = "Epson LQ850"

''Dim P As Object
For Each p In Printers

''    If P.Port = "\\Cms\hp" And P.DeviceName = "HP OfficeJet G Series Printer" And P.DriverName = "hpodj907" And P.TrackDefault = False Then
''
''        Set Printer = P
''        'Printer.TrackDefault = True
''        Exit For
''    End If
''    Next P
''    If Printer.TrackDefault = True Then
''    Printer.TrackDefault = False
''    ElseIf Printer.TrackDefault = False And P.DeviceName = "HP OfficeJet G Series Printer" Then
''    P.TrackDefault = True
''    Else
''
''    End If
''

'IPCR.PrinterName = "Wipro LQ-1050+DX"
'ipcr.WindowState = crptMaximized
'ipcr.Destination = crptToWindow

ipcr.Destination = crptToPrinter
'ipcr.Action = 1
PrintResult% = ipcr.PrintReport
'ipcr.Action = 1
Exit Sub
If PrintResult% <> 0 Then

MsgBox Error(Err)
End If
Exit For
Next
Else
'MsgBox "SERVICE AMOUNT IS NOT EQUAL TO TOTAL AMOUNT"
'PrintResult% = 20515
Exit Sub
End If

End Sub

Sub doc_code()
'DOCNOTFOUND = True
'On Error Resume Next
'    '''Set rs3 = rdb.OpenResultset("select inch_name from MA1031 where ud_name = '" & txt_unit_doc_name & "'", rdOpenKeyset)
'    'Set rs3 = rdb.OpenResultset("select inch_name from MA1031 where ud_name LIKE '" & "A K%" & "'", rdOpenKeyset)
'    '''Set rs33 = rdb.OpenResultset("select doc_code from MA102 where doc_name = '" & rs3("inch_name") & "'", rdOpenKeyset)
'    If rs33flag = False Then
'        Doc_ViewQr(0) = Trim(txt_unit_doc_name)
'        Set rs33 = Doc_ViewQr.OpenResultset(1, 3)
'        rs33flag = True
'    Else
'        Doc_ViewQr(0) = Trim(txt_unit_doc_name)
'        rs33.MoreResults
'        rs33.Requery
'    End If
''    If rs33.RowCount <> 0 Then
''        DOCNOTFOUND = False
''    End If
'
''    MsgBox rs33("DOC_CODE")
''    MsgBox txt_unit_doc_name
'    ''''Set rs33 = rdb.OpenResultset("SELECT doc_code FROM DOCVIEW WHERE ud_name = '" & Trim(txt_unit_doc_name) & "'", rdOpenKeyset)
'    'Set rs33 = rdb.OpenResultset("select doc_code from MA102 where doc_name LIKE '" & "%A K%" & "'", rdOpenKeyset)
'End SubSub doc_code()
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
 If ((rs33.RowCount) = 0) And (Trim(RS33QR(0)) <> "NO UNIT") And (lbl_cat <> "C" Or lbl_cat <> "NC") Then  'GMD ADD FROM ON 21062000
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
Dim opt As String
Dim bal_amt As Long
Dim TEMPREQ As String
Dim trFlag As Boolean

Dim SER_AMT As Long '''A221199
Dim D1 As String
Dim str As String
Dim MNO As String
Dim AUTOMEMO As String

If AmtNotEq = 1 Then ''''2541999
    MsgBox " Memo can't be recorded, try again ! ", vbCritical
    AmtNotEq = 0
    Exit Sub
End If ''''2541999
bal_amt = 0

'rdb.BeginTrans
flag = 1
'Rajesh$&*&*^^%$&*$^^%$#^%^%$%^%^$^%
If flag = 1 Then
'AUTO GENERATION OF MEMO NO.
me_mno.Mask = ">?#######"
S = Right(me_date.Text, 4)
If RS_SERIESflag = False Then
    RQ_SERIES(0) = "A" & S
    Set RS_SERIES = RQ_SERIES.OpenResultset(2, 3)
    RS_SERIESflag = True
Else
   RQ_SERIES(0) = "A" & S
   RS_SERIES.MoreResults
   RS_SERIES.Requery
End If

'*** start modified by v.s.k. on 04012012

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
'AUTOMEMO = "M" & Right(RS_SERIES("YEAR1"), 1) & (Trim(RS_SERIES("SERIES")) + 1)
AUTOMEMO = "M" & Right(RS_SERIES("YEAR1"), 1) & K
If Len(Trim(Mid(AUTOMEMO, 3))) <> 0 Then me_mno = padzero(6, "M", K, 1)
'*** end

' If IsNull(RS_SERIES("MEMO_SRNO")) Then
'     me_mno = RS_SERIES("CHAR_COR_YEAR") & "" & "1"
'    Else
'   me_mno.Text = RS_SERIES("CHAR_COR_YEAR") & "O" & (RS_SERIES("MEMO_SRNO") + 1)
' End If
'Else
   'AUTOMEMO = "M" & Right(me_date.Text, 1) & "1"
   
'End If
'****

'If RS_SERIES.RowCount > 0 Then
'    K = Mid(AUTOMEMO, 3, 8) ' ADDED 11/11/2001
'    k1 = K
'    If Len(Trim(Mid(AUTOMEMO, 3))) <> 0 Then
'        '    me_mno = pad(6, "M", me_mno, 1)
'
'        If RS_SERIES.RowCount <> 0 Then   ' ADDED 11/11/2001
'            me_mno = padzero(6, "M", K, 1)  ' ADDED 11/11/2001
'        Else                              ' ADDED 11/11/2001
'            me_mno = padzero(6, "M", K, 1)
'        End If
'
'    End If
'
''****
'    rq_updtSERIES(0) = k1
'    rq_updtSERIES(1) = "A" & S
'    rq_updtSERIES.Execute
' End If
'AUTO GENERATION OF MEMO NO



rdb.BeginTrans

D1 = Format(me_date.Text, "dd/mm/yyyy")
    D1 = GET_REVDATE(D1)
    
  ''' SANDHYA PROTOCOL PATIENT STARTS
    If Val(nettot) > 0 Then
    'If CDbl(Format(nettot, "########.00")) > 0 Then
   ''' PROTOCOL PATIENT ENDS
    Ta107_InsrtQr(0) = me_mno.FormattedText 'rs4("memo_no")
    Ta107_InsrtQr(1) = "O" 'rs4("io_memo")
    Ta107_InsrtQr(2) = D1 'Format(me_date.Text, "dd/mm/yyyy") 'rs4("memo_date")
    Ta107_InsrtQr(3) = UCase(lbl_dname) 'rs4("dept_code")
    Ta107_InsrtQr(4) = CASENO 'me_cno 'rs4("case_no")
    If Len(Trim(cbo_reqno.Text)) <> 0 Then 'add on 24/7
    Ta107_InsrtQr(5) = Trim(lbl_cat) 'Trim(lbl_cat) 'rs4("MEMO_CATG") 'add on 24/7
    Else 'add on 24/7
    Ta107_InsrtQr(5) = Trim(lbl_cat) 'add on 24/7
    End If 'add on 24/7
    
    'COMMENTED ASND ADDED ON 24/01/2011 FOR STORING USERID
    Ta107_InsrtQr(6) = usr_name
'  If Len(Trim(txt_doc_sec_name)) = 0 Then
'  Ta107_InsrtQr(6) = "userid"
'  Else
'Ta107_InsrtQr(6) = txt_doc_sec_name
'  End If
    Ta107_InsrtQr(7) = ipadd & " - " & time1  'DIS" 'txt_rem 'rs4("rem")
'Ta107_InsrtQr(11) = " " 'Trim(cbo_reqno.Text) 'txt_req_no 'comm on 24/7
'  Ta107_InsrtQr(12) = Null ' comm on 24/7
'  If me_rdate.Text <> "" Then
'    ass_d1 = Left(me_rdate.Text, 2) & "/" & Mid(me_rdate.Text, 3, 2) & "/" & Right(me_rdate.Text, 4) 'Format(me_rdate.Text, "dd/mm/yyyy")
'    ass_d1 = GET_REVDATE(ass_d1)
'    Ta107_InsrtQr(12) = ass_d1
'  End If
'comm on 24/7

    Call dcode
    
    If rs2("pat_reg_ty") = "G" Then
        Ta107_InsrtQr(8) = Val(" ") 'rs4("doc_code")
    ElseIf rs2("pat_reg_ty") = "P" Then
        
    If rs33flag = False Then
        Doc_ViewQr(0) = Trim(txt_unit_doc_name)
        Set rs33 = Doc_ViewQr.OpenResultset(1, 3)
        rs33flag = True
    Else
        Doc_ViewQr(0) = Trim(txt_unit_doc_name)
        rs33.MoreResults
        rs33.Requery
    End If
       If rs33.RowCount > 0 Then
            Ta107_InsrtQr(8) = rs33("doc_code") 'rs4("doc_code")
        Else
            Ta107_InsrtQr(8) = 0
        End If
    ElseIf rs2("pat_reg_ty") = "R" Then
         If rs2("U_D") = "D" Then
                If rs33flag = False Then
                    Doc_ViewQr(0) = Trim(txt_unit_doc_name)
                    Set rs33 = Doc_ViewQr.OpenResultset(1, 3)
                    rs33flag = True
                Else
                    Doc_ViewQr(0) = Trim(txt_unit_doc_name.Text)
                    rs33.MoreResults
                    rs33.Requery
                End If
                If rs33.RowCount > 0 Then
                        Ta107_InsrtQr(8) = rs33("doc_code")
                End If
          ElseIf rs2("U_D") = "U" Then
                     
             Ta107_InsrtQr(8) = Val(" ")
         End If
    End If
    
    If rs2("PAT_TYPE") = "CR" And BPPATFLAG = False Then
        Ta107_InsrtQr(9) = "C" 'rs4("amt_rcvd")
   '''' ADDED BY SANDHYA DATED 7.7.2012 FOR BP RGJAY PATIENT
    ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = True Then
        Ta107_InsrtQr(9) = "R"
        
    ElseIf rs2("PAT_TYPE") = "TA" Then
        Ta107_InsrtQr(9) = "T" 'rs4("amt_rcvd")
    'V001 PATIENT DEPOSIT STARTS
    ElseIf rs2("PAT_TYPE") = "CP" And PATDEPO = True Then
        Ta107_InsrtQr(9) = "P"
    'V001 PATIENT DEPOSIT ENDS
     '''ADDED BY CHARULATA FOR SMART CARD'''''''
    ElseIf rs2("PAT_TYPE") = "CP" And scard = True Then
        Ta107_InsrtQr(9) = "S"
        ''''UPTO HERE''''
    Else
        Ta107_InsrtQr(9) = Null 'rs4("amt_rcvd")
        
    End If
    If PROTFLG = True Then
     Ta107_InsrtQr(10) = CDbl(Format(nettot, "#,##,##0.00"))
    Else
    Ta107_InsrtQr(10) = CDbl(Format(lbl_tot_amt, "#,##,##0.00")) 'rs4("memo_amt")
    End If
    Ta107_InsrtQr(11) = "N"
    Ta107_InsrtQr.Execute 'insert Ta107

  End If
'''ADDED BY CHARULATA
If Trim(rs2("PAT_TYPE")) = "CP" And scard = True Then
    If updatesmart = 1 Then
         UpdateTransaction
            DisplayPole ("Thank You " & Trim(me_cno.FormattedText) & "  Balance " & PadSpace(12, lbl_deposit.Caption))
        
    Else
        MsgBox ("SMART ACOOUNT NOT UPDATED.")
        GoTo MSG
    End If
End If
''''UPTO HERE



 '  If nettot > 0 Then
    For iRow = 0 To mTotalRows&
        If Len(Trim(UserData(0, iRow))) > 0 Then
            If Not IsEmpty(UserData(0, iRow)) Then
            
            '' SANDHYA PROTOCOL PATIENT STARTS
             If Val(nettot) > 0 Then
             
               'PROTSERV_QR(0) = PROTCASE_RS("IRBno")
               
               PROTSERV_QR(0) = MIRBNO
               PROTSERV_QR(1) = Left(UserData(0, iRow), 4)
'               PROTSERV_QR(2) = ZAS400DTTM!DT1
'               PROTSERV_QR(3) = ZAS400DTTM!DT1
               Set PROTSERV_RS = PROTSERV_QR.OpenResultset(2, 3)
               PROTSERV_RS.Requery
               If PROTSERV_RS.RowCount = 0 Then
                  
                  Ta1071_InsrtQr(0) = me_mno.FormattedText 'rs7("memo_no")
                  Ta1071_InsrtQr(1) = Left(UserData(0, iRow), 4) 'rs7("ser_code")
                    Ta1071_InsrtQr(2) = UserData(2, iRow) 'rs7("qty")
                    Ta1071_InsrtQr(3) = Val(UserData(3, iRow)) 'rs7("ser_amt")
    
                    If Val(UserData(4, iRow)) <> 0 Then
                        Ta1071_InsrtQr(4) = Val(UserData(4, iRow)) 'rs7("doc_code")
                         Else
                       Ta1071_InsrtQr(4) = 0  'rs7("doc_code")
                    End If
                    'Ta1071_InsrtQr(5) = Mid(UserData(4, iRow), 6, 13)
                    'If Len(Trim(cbo_reqno.Text)) <> 0 Then 'add on 24/7
                    If Len(Trim(UserData(0, iRow))) <> 0 Then
                     If Len(Trim(UserData(0, iRow))) > 4 Then
                     Ta1071_InsrtQr(5) = Mid(UserData(0, iRow), 6, 13) 'ADD ON 24/7
                                        
                     Ta1071_InsrtQr(6) = GET_REVDATE(Mid(UserData(0, iRow), 20, 10)) ' add on 24/7
                     Ta1071_InsrtQr(7) = Mid(UserData(0, iRow), 31) ' add on 24/7
                     
'                    Ta1071_InsrtQr(5) = Mid(UserData(0, iRow), 6, 13) 'ADD ON 24/7
'                    Ta1071_InsrtQr(6) = Mid(UserData(0, iRow), 20, 10) ' add on 24/7
'                    Ta1071_InsrtQr(7) = Mid(UserData(0, iRow), 31) ' add on 24/7
                    Else  'add on 24/7
                        Ta1071_InsrtQr(5) = "" 'Mid(UserData(0, iRow), 6, 13) 'ADD ON 24/7
                        Ta1071_InsrtQr(6) = D1 ' CDate(Mid(UserData(0, iRow), 15)) ' add on 24/7
                        '''' ADDED BY SANDHYA DATED 7.7.2012 FOR BP PATIENT RGJAYOJANA
                       If BPPATFLAG = True Then
                           Ta1071_InsrtQr(7) = "BP"
                       Else
                           Ta1071_InsrtQr(7) = ""
                       End If
                   
                        
                    End If
                    
'                        Ta1071_InsrtQr(5) = "" 'Mid(UserData(0, iRow), 6, 13) 'ADD ON 24/7
'                        Ta1071_InsrtQr(6) = Null ' CDate(Mid(UserData(0, iRow), 15)) ' add on 24/7
'                        Ta1071_InsrtQr(7) = ""
                    
                   End If  'add on 24/7
                               
                   
                    '*****
                    '17/10/2002-Parag/Sandhya
                    'Bifurcating Ser_amt as Prof_amt and Hosp_amt.
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
                        Ta1071_InsrtQr(8) = UserData(2, iRow) * zrsgetservice!a_prof_ch
                        Ta1071_InsrtQr(9) = UserData(2, iRow) * (zrsgetservice!a_hosp_ch - zrsgetservice!a_prof_ch)
                        Ta1071_InsrtQr(10) = "X"
                         Case "B"
                        Ta1071_InsrtQr(8) = UserData(2, iRow) * zrsgetservice!b_prof_ch
                        Ta1071_InsrtQr(9) = UserData(2, iRow) * (zrsgetservice!b_hosp_ch - zrsgetservice!b_prof_ch)
                        Ta1071_InsrtQr(10) = "X"
                         Case "F"
                        Ta1071_InsrtQr(8) = UserData(2, iRow) * zrsgetservice!fn_prof_ch
                        Ta1071_InsrtQr(9) = UserData(2, iRow) * (zrsgetservice!fn_hosp_ch - zrsgetservice!fn_prof_ch)
                        Ta1071_InsrtQr(10) = "X"
                         Case "D"
                        Ta1071_InsrtQr(8) = UserData(2, iRow) * zrsgetservice!d_prof_ch
                        Ta1071_InsrtQr(9) = UserData(2, iRow) * (zrsgetservice!d_hosp_ch - zrsgetservice!d_prof_ch)
                        Ta1071_InsrtQr(10) = "X"
                         Case "C"
                        Ta1071_InsrtQr(8) = 0
                        Ta1071_InsrtQr(9) = UserData(2, iRow) * zrsgetservice!c_ch
                        Ta1071_InsrtQr(10) = "X"
                         Case "BP"
                        Ta1071_InsrtQr(8) = 0
                        Ta1071_InsrtQr(9) = UserData(2, iRow) * zrsgetservice!c_ch
                        Ta1071_InsrtQr(10) = "X"
                         Case Else   ' "NC" case
                        Ta1071_InsrtQr(8) = 0
                        Ta1071_InsrtQr(9) = UserData(2, iRow) * zrsgetservice!nc_ch
                        Ta1071_InsrtQr(10) = "X"
                         End Select
                        Else
                            Ta1071_InsrtQr(8) = 0
                            Ta1071_InsrtQr(9) = CDbl(UserData(3, iRow))
                            Ta1071_InsrtQr(10) = "X"
                        End If

                    End If
                    ''*******on 26/06/03
                    Ta1071_InsrtQr(11) = CASENO
                    '*****
                    Ta1071_InsrtQr.Execute
                 '' PROTOCOL
                 End If
                 '' PROTOCOL PATIENT ENDS
               End If
               End If
           End If
         If Len(Trim(UserData(0, iRow))) > 4 Then
             'If Val(nettot) > 0 Then
                    'add on 24/7
                    rq_updtreqdet(0) = D1
                    rq_updtreqdet(1) = Trim(me_mno.FormattedText) '"M-" & me_mno.Text
                    '' checking for icu req and making E to bypass exam over
                    'made P from E
                        
                    If pathology.lbl_icureq.Visible = True Or Trim(pathology.txt_status.Text) = "IP" Then
                        rq_updtreqdet(2) = "P"
                    Else
                           rq_updtreqdet(2) = "M"
                    End If
                    
                    '***** Shraddha 11/10/2011 ******
                    If Left(Trim(UserData(0, iRow)), 4) = "G402" And Trim(pathology.txt_catgory.Text) = "NC" Then
                        rq_updtreqdet(2) = "M"
                    End If
                    '*********************************

                   '***** Shraddha 27/07/10 *****
                   SELDISWARD_QR(0) = Trim(ipadd)
                   Set SELDISWARD_RS = SELDISWARD_QR.OpenResultset(1, 3)
                   
                   If Not SELDISWARD_RS.EOF And Trim(pathology.txt_modcode) = "EL" Then
                        If Trim(pathology.txt_status.Text) = "OP" Then
                            rq_updtreqdet(2) = "P"
                        End If
                   End If
                   '*****************************
                   
                   
                   '***** Added by Shraddha on 27/10/12 *****
                   'To make requisition status P for M% and I%
                   SEL_DISICU_PM_QR(0) = Trim(ipadd)
                   Set SEL_DISICU_PM_RS = SEL_DISICU_PM_QR.OpenResultset(1, 3)
                   
                   If Not SEL_DISICU_PM_RS.EOF Then
                        rq_updtreqdet(2) = "P"
                   Else
                        rq_updtreqdet(2) = "M"
                   End If
                   
                   '*****************************************
                   
                   
                    rq_updtreqdet(3) = Trim(CASENO)
                     rq_updtreqdet(4) = Trim(Mid(UserData(0, iRow), 6, 13)) 'Trim(cbo_reqno.Text) 'txt_req_no.Text ' add on 24/7
                    'rq_updtreqdet(5) = lbl_dname
                    '  rq_updtreqdet(6) = ""
                    rq_updtreqdet.Execute
                '----- IN PATHOLOGY VPP  CASE BY PASSING THE FLAG
            If Len(Mid(UserData(0, iRow), 6, 13)) <> 0 Then
                    If SELRREQRSFLAG = False Then
                         SELRREQQR(0) = CASENO
                         SELRREQQR(1) = (Mid(UserData(0, iRow), 6, 13))
                         Set SELRREQRS = SELRREQQR.OpenResultset(1, 3)
                         SELRREQRSFLAG = True
                    Else
                         SELRREQQR(0) = CASENO
                         SELRREQQR(1) = (Mid(UserData(0, iRow), 6, 13))
                         SELRREQRS.MoreResults
                         SELRREQRS.Requery
                    End If
                If SELRREQRS.RowCount > 0 Then
                   If SELRREQRS("RECORDNO") = 1 Then
                        REQUISITN_UPDT(0) = "VPP"
                        REQUISITN_UPDT(1) = me_date.FormattedText
                        REQUISITN_UPDT(2) = "P"
                        REQUISITN_UPDT(3) = Mid(UserData(0, iRow), 6, 13)
                        REQUISITN_UPDT.Execute
                   End If
                End If
            End If
        '-----IN PATHOLOGY VPP OR LYMPHOMA REGISTRY CASE BY PASSING THE FLAG
        
         ' End If
         End If
  '------
        
If rs2("PAT_TYPE") = "TA" Then

        If Len(Mid(UserData(0, iRow), 6, 13)) <> 0 Then
        
          REQUISITN_UPDT(0) = "TRUST"
          REQUISITN_UPDT(1) = me_date.FormattedText
          REQUISITN_UPDT(2) = "P"
          REQUISITN_UPDT(3) = Mid(UserData(0, iRow), 6, 13)
          REQUISITN_UPDT.Execute
         
       
        UPDTA107_QR(0) = "T"
        UPDTA107_QR(1) = me_mno.FormattedText
        UPDTA107_QR.Execute
        End If
End If

''''''''''added by charulata for smartcard
If rs2("PAT_TYPE") = "CP" And scard = True Then

        If Len(Mid(UserData(0, iRow), 6, 13)) <> 0 Then
        
          REQUISITN_UPDT(0) = "SCARD"
          REQUISITN_UPDT(1) = me_date.FormattedText
          REQUISITN_UPDT(2) = "P"
          REQUISITN_UPDT(3) = Mid(UserData(0, iRow), 6, 13)
          REQUISITN_UPDT.Execute
             
        
        End If
End If

''''upto here

''   icu cases
'SAVING REQ NO FOR ICU CASES

TEMPREQ = Mid(UserData(0, iRow), 6, 13)
 ' If pathology.lbl_icureq.Visible = True Then
 '       disreq_updt(0) = "E"
 '       disreq_updt(1) = Mid(UserData(0, iRow), 6, 13)
 ' End If
  
''   icu cases

  '----
        If rs2("PAT_TYPE") = "CR" And BPPATFLAG = False Then

           If Len(Mid(UserData(0, iRow), 6, 13)) <> 0 Then
            REQUISITN_UPDT(0) = "COMPANY"
            REQUISITN_UPDT(1) = me_date.FormattedText
            REQUISITN_UPDT(2) = "P"
            REQUISITN_UPDT(3) = Mid(UserData(0, iRow), 6, 13)
            REQUISITN_UPDT.Execute
          
         
       
            UPDTA107_QR(0) = "C"
            UPDTA107_QR(1) = me_mno.FormattedText
            UPDTA107_QR.Execute
            End If
        End If

    'V001 PATIENT DEPOSIT STARTS
    
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
    
   
      If rs2("PAT_TYPE") = "CP" And PATDEPO = True Then
          If Len(Mid(UserData(0, iRow), 6, 13)) <> 0 Then
            REQUISITN_UPDT(0) = "DEPOSIT"
            REQUISITN_UPDT(1) = me_date.FormattedText
            REQUISITN_UPDT(2) = "P"
            REQUISITN_UPDT(3) = Mid(UserData(0, iRow), 6, 13)
            REQUISITN_UPDT.Execute
            
            UPDTA107_QR(0) = "P"
            UPDTA107_QR(1) = me_mno.FormattedText
            UPDTA107_QR.Execute
            
          End If
      End If
      
     'V001 PATIENT DEPOSIT ENDS
                    
                    
   'ADDED BY SANDHYA PROTOCOL PATIENT STARTS
   
      If PROTFLG = True Then
          If Len(Mid(UserData(0, iRow), 6, 13)) <> 0 Then
            REQUISITN_UPDT(0) = "PROTOCOL"
            REQUISITN_UPDT(1) = me_date.FormattedText
            REQUISITN_UPDT(2) = "P"
            REQUISITN_UPDT(3) = Mid(UserData(0, iRow), 6, 13)
            REQUISITN_UPDT.Execute
            
'            UPDTA107_QR(0) = "P"
'            UPDTA107_QR(1) = me_mno.FormattedText
'            UPDTA107_QR.Execute
            
          End If
      End If
      
     'V001 PROTOCOL PATIENT ENDS
   '----
  '------
                    'add on 24/7
                    
       ''''' ADDED BY SANDHYA DATED 11.6.2012 FOR POSTING DATA IN TABLE ACTTMHMEMO
  ''''' COMMON REQUISITION GENERATED WHEN PATIEN RESIDES AT ACTREC & REPORT
  ''''' GENERATED AT TMH

  
  If ACTtmhCn = True Then
   If Left(ipadd, 6) = "10.100" Then
    If Len(Mid(UserData(0, iRow), 6, 13)) <> 0 Then
      actmemflg = True
      Call ACTREQMAIN
   ' str = "insert into patoolslib.acttmhmemo(memo_no,memo_date,ser_code,ser_amt,qty,transtat,ipaddrs,userid,recstat) values (?,?,?,?,?,?,?,?,?)"
        INSACTMHMEMO(0) = me_mno.FormattedText
        INSACTMHMEMO(1) = me_date.FormattedText
        INSACTMHMEMO(2) = Left(UserData(0, iRow), 4)
        INSACTMHMEMO(3) = Val(UserData(3, iRow))
        INSACTMHMEMO(4) = UserData(2, iRow)
        INSACTMHMEMO(5) = ""
        INSACTMHMEMO(6) = ipadd
        INSACTMHMEMO(7) = usr_name
        INSACTMHMEMO(8) = ""
        INSACTMHMEMO.Execute
    
        
        
      If rs2("PAT_TYPE") = "CR" And BPPATFLAG = False Then
        UPDACTTMHREQ(0) = "P"
        UPDACTTMHREQ(1) = "COMPANY" '
        UPDACTTMHREQ(2) = me_date.FormattedText
      ElseIf rs2("PAT_TYPE") = "TA" Then
        UPDACTTMHREQ(0) = "P"
        UPDACTTMHREQ(1) = "TRUST"
         UPDACTTMHREQ(2) = me_date.FormattedText
      ElseIf rs2("PAT_TYPE") = "CP" Then
        UPDACTTMHREQ(0) = "M"
        UPDACTTMHREQ(1) = Null
         UPDACTTMHREQ(2) = me_date.FormattedText
      End If
        UPDACTTMHREQ(3) = me_mno.FormattedText
        UPDACTTMHREQ(4) = me_date.FormattedText
        UPDACTTMHREQ(5) = CDbl(Format(lbl_tot_amt, "#,##,##0.00"))
        UPDACTTMHREQ(6) = Mid(UserData(0, iRow), 6, 13)
        UPDACTTMHREQ.Execute
    Else
       actmemflg = False
    End If
    
   End If
 End If
    
    Next iRow
 End If
'End If ' nettot > 0

'' SANDHYA DATED 11.8.2008

'' TO INSERT INTO PROTOCOL MEMO TABLE
If Val(prottot) > 0 Then
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

PROTMEMO_Qr(9) = prottot
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
PROTMEMO_Qr(16) = "O"
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
'' PROTOCOL PATIENT ENDS

If rs2("PAT_TYPE") = "TA" Then
    If FLGTA = 1 And PROTFLG = False Then
'''         Set rs_tr = rdb.OpenResultset("select * from ta113 where case_no = '" & CASENO & "'", 2, 3)
'''         rs_tr.Edit
'''         rs_tr("ser_amt") = rs_tr("ser_amt") + lbl_tot_amt
'''         rs_tr.update
         Ta113_UpdtQr(0) = CDbl(lbl_tot_amt)
         Ta113_UpdtQr(1) = CASENO
         Ta113_UpdtQr.Execute 'update ta113
         ''''rdb.Execute ("update ta113 set ser_amt = ser_amt + " & lbl_tot_amt & " where case_no = '" & CASENO & "'")
    End If
ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = False Then
        If FLGCR = 1 Then
'''            Set rs_cr = rdb.OpenResultset("select * from ta114 where case_no = '" & CASENO & "'", 2, 3)
'''            rs_cr.Edit
'''            rs_cr("ser_amt") = rs_cr("ser_amt") + lbl_tot_amt
'''            rs_cr.update
            Ta114_UpdtQr(0) = CDbl(lbl_tot_amt)
            Ta114_UpdtQr(1) = CASENO
            Ta114_UpdtQr.Execute
            ''''rdb.Execute ("update ta114 set ser_amt = ser_amt + " & lbl_tot_amt & " where case_no = '" & CASENO & "'")
        End If
''''' ADDED BY SANDHYA DATED 7.7.2012 FOR BP RGJAY PATIENT
ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = True Then
     UPDbppatn(0) = CDbl(lbl_tot_amt.Caption)
     UPDbppatn(1) = CDbl(lbl_tot_amt.Caption)
     UPDbppatn(2) = CASENO
     UPDbppatn.Execute
     
     Set RSBPPATN = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.BPPATIENT WHERE CASENO ='" & CASENO & "'", 1, 1)
     
     INS_BPPATMEMO(0) = CASENO
     INS_BPPATMEMO(1) = "BP"
     INS_BPPATMEMO(2) = me_mno.FormattedText
     INS_BPPATMEMO(3) = me_date.FormattedText
     INS_BPPATMEMO(4) = CDbl(lbl_tot_amt)
     INS_BPPATMEMO(5) = Trim(RSBPPATN("SCHEM_CODE"))
     INS_BPPATMEMO(6) = RSBPPATN("SCHEM_AMT")
     INS_BPPATMEMO(7) = usr_name
     INS_BPPATMEMO(8) = Winsock1.LocalIP
     INS_BPPATMEMO(9) = "O"
     INS_BPPATMEMO.Execute
        
'V001 PATIENT DEPOSIT STARTS
ElseIf rs2("PAT_TYPE") = "CP" Then
        If PATDEPO = True Then
            Ta116_UpdtQr(0) = CDbl(lbl_tot_amt.Caption)
            Ta116_UpdtQr(1) = CASENO
            Ta116_UpdtQr.Execute
End If
'V001 PATIENT DEPOSIT ENDS

''' ADDED BY SANDHYA PROTOCOL PATIENT DATED 01.10.2008

If PTOT > 0 Then


    PROTCASE_UPD(0) = CDbl(prottot)
    PROTCASE_UPD(1) = CASENO
    PROTCASE_UPD(2) = Trim(PROTCASE_RS("IRBNO"))
    PROTCASE_UPD.Execute
            
End If
''' END OF PROTOCOL


End If


rdb.CommitTrans

'-------------
rdb.BeginTrans

'' SANDHYA PROTOCOL PATIENT STARTS
If PROTFLG = False Then
    MAPING
Else
    protMAP
End If
'' PROTOCOL PATIENT ENDS
rdb.CommitTrans


Exit Sub

MSG:
rdb.RollbackTrans
MsgBox Error(Err)
MsgBox "Memo not saved!", vbCritical
'MsgBox "refresh1 completed" 'remove it
End Sub
Public Sub protMAP()
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
                    zsercodemap(1) = Left(Trim(RS_PRTMEMO!ser_code), 4)
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
'Dim CHKAMT As Integer
Dim CHKAMT As Double
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
'''ipcr.ReportFileName = "C:\PB121298\OPMEMO.RPT"
'''ipcr.Action = 1

ipcr.ReportFileName = "C:\PB121298\OPROTMEMO.rpt"
ipcr.SelectionFormula = "{protmemo.MEMO_NO} = '" & me_mno.Text & "'"
' For J = 0 To 130 '261099
'    ipcr.Formulas(J) = " "
'Next J

If flag = 1 Then
    ipcr.Formulas(0) = "fdeptname = '" & Com_dno.Text & "'"
    'ipcr.Formulas(2) = "fcaseno = '" & CASENO & "'"
    Exist_ChkQr(0) = me_mno.Text    '"M-" & me_mno.Text
    str12 = CASENO
    Set rs7 = Exist_ChkQr.OpenResultset(1, 1)
    STR10 = "*A*"
Else
    ipcr.Formulas(0) = "fdeptname = '" & lbl_dno & "'"
    Exist_ChkQr(0) = M_number '"M-" & M_number
    str12 = lbl_cno
    Set rs7 = Exist_ChkQr.OpenResultset(1, 1)
    STR10 = "*V*"
End If

'If Label12.Caption = "DOCTOR NAME : " Then
'     ipcr.Formulas(1) = "doc_name = " & a & txt_unit_doc_name & a & ""
'     ipcr.Formulas(2) = "doc_code = " & a & lbl_dcode & a & ""
' ElseIf Label12.Caption = "UNIT NAME : " Then
'     ipcr.Formulas(1) = "doc_name = " & a & txt_unit_doc_name & a & ""
'     ipcr.Formulas(2) = "doc_code = " & a & 0 & a & ""
' Else
     ipcr.Formulas(1) = "doc_name = " & a & txt_unit_doc_name & a & ""
     ipcr.Formulas(2) = "doc_code = " & a & 0 & a & ""
'End If
''IPCR.Formulas(2) = "Fcaseno = '" & CASENO & "'"
'
''ipcr.Formulas(3) = "fdcode = " & a & lbl_dname & a & ""
ipcr.Formulas(4) = "NAME_CON = " & a & Trim(UCase(LBL_NAME)) & a & ""
ipcr.Formulas(5) = "ftot = " & a & " *****" & Format(lbl_tot_amt, "#,##,##,##,##0.00") & a & ""
ipcr.Formulas(6) = "Ptot = " & a & " *****" & Format(prottot, "#,##,##,##,##0.00") & a & ""
ipcr.Formulas(7) = "nettot = " & a & " *****" & Format(nettot, "#,##,##,##,##0.00") & a & ""
'
'
'STR1 = RTrim(username) 'change this var
 STR5 = RTrim(usr_name)
 STR6 = Format(DATE11, "DD/MM/YYYY")
 STR7 = time1
 'Exist_ChkQr(0) = "M-100010"
'Set rs22_CHK = Exist_ChkQr.OpenResultset(1, 1)
If rs7.RowCount <> 0 Then
    STR9 = Trim(rs7("D_S_NAME"))
End If
 'STR8 = Trim(STR5) + "  " + Trim(STR6) + "  " + Trim(STR7)
 STR8 = Trim(STR9) + " / " + Trim(STR5) + " " + Trim(STR10) + " " + " " + STR11 + " " + Trim(STR6) + "  " + Trim(STR7)
ipcr.Formulas(8) = "us_dt_tm = " & a & STR8 & a & ""
'
 If rs2("PAT_TYPE") = "TA" Then
     ipcr.Formulas(9) = "cr_tr = '" & rs_tr("TR_ACC_NO") & "'"
     ipcr.Formulas(10) = "PATTP = '" & "TRUST PATIENT" & "'"
 ElseIf rs2("PAT_TYPE") = "CR" Then
    ipcr.Formulas(9) = "cr_tr = '" & Trim(rs_cr("CR_CARD")) & "'"
 ElseIf rs2("PAT_TYPE") = "CP" And PATDEPO = True Then
    ipcr.Formulas(9) = "cr_tr = '" & rs_tr("PD_ACC_NO") & "'"
    ipcr.Formulas(10) = "PATTP = '" & "DEPOSIT PATIENT" & "'"
 ElseIf PROTFLG = True Then
    ipcr.Formulas(9) = "cr_tr = '" & "    " & "'"
    ipcr.Formulas(10) = "PATTP = '" & "PROTOCOL PATIENT" & "'"
 Else
    ipcr.Formulas(9) = "cr_tr = '" & "    " & "'"
    ipcr.Formulas(10) = "PATTP = '" & " " & "'"
 End If
'***** 09/02/10 *****
ipcr.Formulas(11) = "PINAME = '" & PINAME & "'"
'********************
    'DBGrid1.Refresh




ipcr.CopiesToPrinter = 1
  
''Dim P As Object
For Each p In Printers


'IPCR.PrinterName = "Wipro LQ-1050+DX"
'ipcr.WindowState = crptMaximized
ipcr.Destination = crptPRINTER
'PrintResult% = ipcr.PrintReport
ipcr.Action = 1
If PrintResult% <> 0 Then
MsgBox Error(Err)
End If
Exit For
Next

End Sub
Sub add_edit()
flag = 1
''If AmtNotEq = 1 Then
''    MsgBox " Memo can't be recorded, try again ! ", vbCritical
''    AmtNotEq = 0
''    Exit Sub
''End If
''
''If flag = 1 Then
''       '''rs4.AddNew
''       '''rs7.AddNew
''Else
''       '''rs4.Edit
''       '''rs7.Edit
''End If
'''MsgBox "in add_edit, after this assign" 'remove it
''        assign
'*****
Dim mser_amt As Long
mser_amt = 0
If AmtNotEq = 1 Then
    MsgBox " Memo can't be recorded, try again ! ", vbCritical
    AmtNotEq = 0
    Exit Sub
End If

If flag = 1 Then '[NEW]
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
       
       If CDbl(mser_amt) <> CDbl(lbl_tot_amt) Then
          MsgBox " Total Memo Amount Does Not Match ! ", vbCritical
          Exit Sub
       Else
          assign
       End If
      
End If



'*****
End Sub

Private Sub Back_Click()
'Dim r1 As rdoResultset
'For Each r1 In rdb.rdoResultsets
'
'r1.Close
'Next
REQGOTFALG = False
ReDim UserData(0 To 5, 0 To 12)
DBGrid1.Refresh
cleartxt
Me.Hide

'MENU1.Show
'tmh_ris_mnu.Show
'SELMNU.Show '**30/6/99

End Sub

Private Sub cbo_reqno_GotFocus()
Dim REQNO As String
'24-AUG-09 SHWETA
CASENO = pathology.msk_caseno.FormattedText  'me_cno.FormattedText
'''added by charulata
If Trim(pathology.PATYPE) = "CP" Then
    Set PCARD = New PISWriteCard.ClsWriteCard

    If GETSMARTCARD(CASENO, PCARD.GetCSNR(Strkey, PINFO)) = "1" Then
        lbl_smartcard.Visible = True
        lbl_deposit.Visible = True
        lbl_deposit.Caption = scard_bal_amt
        ''''ADDED FOR WRITING IN ERROR LOG
        str = PCARD.GetLastTransaction(Strkey, PINFO)
        If Val(PINFO.FinalAmt) <> Val(scard_bal_amt) Then
            Call INSERTERROR("", CASENO, PINFO.FinalAmt, Val(scard_bal_amt))
        End If
        scard = True
        ''''UPTO HERE
    Else
        lbl_smartcard.Visible = False
        
        lbl_deposit.Visible = False
        scard = False
    End If
Set PCARD = Nothing
End If

'upto here

If REQGOTFALG = False Then
btn_vis (1)
AmtNotEq = 0
REQGOTFALG = True
 i = pathology.Data1.Recordset.RecordCount
 
 If i <> 0 Then
 
    pathology.Data1.Recordset.MoveFirst
    '****  filling the reqno combo box
    While Not pathology.Data1.Recordset.EOF
            
            'Change Made For NC IN-PATIENTS
                       
            'added by milind on 01-09-2006
            patndt_qr(0) = CASENO
            patndt_qr(1) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
            Set patndt_rs = patndt_qr.OpenResultset(1, 3)
            
            If patndt_rs.RowCount > 0 Then
            '**********************'24-AUG-09 SHWETA
            lbl_cat.Caption = Trim(UCase(pathology.txt_catgory.Text))
            '**********************
                 If lbl_cat.Caption = "NC" And patndt_rs!patnstat = "IP" Then
                    INNC = True
                 End If
            End If
            'upto here
            
            If Trim(lbl_cat.Caption) = "NC" And INNC = True Then
                     If RS_reqdetflag = False Then
                    rq_reqdet(0) = CASENO
                    rq_reqdet(1) = "IP"
                    rq_reqdet(2) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                    Set RS_reqdet = rq_reqdet.OpenResultset(1, 3)
                    RS_reqdetflag = True
                Else
                    rq_reqdet(0) = CASENO
                    rq_reqdet(1) = "IP"
                    rq_reqdet(2) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                    RS_reqdet.MoreResults
                    RS_reqdet.Requery
                End If
            Else
            ' problem for vpp
                If RS_reqdetflag = False Then
                    rq_reqdet(0) = CASENO
                    rq_reqdet(1) = "OP"
                    rq_reqdet(2) = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
                    Set RS_reqdet = rq_reqdet.OpenResultset(1, 3)
                    RS_reqdetflag = True
                Else
                    rq_reqdet(0) = CASENO
                    rq_reqdet(1) = "OP"
                    rq_reqdet(2) = Trim(req_chglen(pathology.Data1.Recordset.Fields("reqno")))
                    RS_reqdet.MoreResults
                    RS_reqdet.Requery
                End If
            End If
       
 
    '**** filling the reqno combo box
            If RS_reqdet.RowCount > 0 Then
                If Trim(RS_reqdet("ser_code")) = "8888" Then
                    ''************ADD BY MONO FOR BP************
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
            pathology.Data1.Recordset.MoveNext
    Wend
    '  this is done because reqno iis getting printing
    
   pathology.Data1.Recordset.MoveFirst
   
 End If
 ReDim UserData(0 To 5, 0 To 12)
 DBGrid1.Refresh
End If
txt_rem.Enabled = True
If cbo_reqno.Text = "" Then
'cbo_reqno.AddItem pathology.SHW_REQNO
End If
'MsgBox pathology.txt_bcreqno
End Sub

Private Sub cbo_reqno_LostFocus()
Dim REQNO As String
On Error GoTo Errhandler
'CASENO = me_cno.FormattedText
'If Len(Trim(cbo_reqno.Text)) = 0 Then
'cbo_reqno.ListIndex = -1
''added for ris modification on 11052001
'   reqno = Left(cbo_reqno.Text, 3) & Mid(cbo_reqno.Text, 5, 2) & Mid(cbo_reqno.Text, 8, 2) & Right(cbo_reqno.Text, 6)
'    If RS_reqdetflag = False Then
'         rq_reqdet(0) = CASENO
'         rq_reqdet(1) = "OP"
'         rq_reqdet(2) = reqno
'         Set RS_reqdet = rq_reqdet.OpenResultset(1, 3)
'         RS_reqdetflag = True
'    Else
'         rq_reqdet(0) = CASENO
'         rq_reqdet(1) = "OP"
'         rq_reqdet(2) = reqno
'         RS_reqdet.MoreResults
'         RS_reqdet.Requery
'    End If
'
'
'Else
'    reqno = Left(cbo_reqno.Text, 3) & Mid(cbo_reqno.Text, 5, 2) & Mid(cbo_reqno.Text, 8, 2) & Right(cbo_reqno.Text, 6)
'
'    '23/01/2002-Parag
'    'Change Made For NC IN-PATIENTS
'    If Trim(lbl_cat.Caption) = "NC" And INNC = True Then
'        If RS_reqdetflag = False Then
'            rq_reqdet(0) = CASENO
'            rq_reqdet(1) = "IP"
'            rq_reqdet(2) = reqno
'            Set RS_reqdet = rq_reqdet.OpenResultset(1, 3)
'            RS_reqdetflag = True
'        Else
'            rq_reqdet(0) = CASENO
'            rq_reqdet(1) = "IP"
'            rq_reqdet(2) = reqno
'            RS_reqdet.MoreResults
'            RS_reqdet.Requery
'        End If
'    Else
'        If RS_reqdetflag = False Then
'            rq_reqdet(0) = CASENO
'            rq_reqdet(1) = "OP"
'            rq_reqdet(2) = reqno
'            Set RS_reqdet = rq_reqdet.OpenResultset(1, 3)
'            RS_reqdetflag = True
'        Else
'            rq_reqdet(0) = CASENO
'            rq_reqdet(1) = "OP"
'            rq_reqdet(2) = reqno
'            RS_reqdet.MoreResults
'            RS_reqdet.Requery
'        End If
'    End If
    'rq_reqdet(3) = lbl_dname
  '   rq_reqdet(4) = ""
    'Set RS_reqdet = rq_reqdet.OpenResultset(1, 3)
    '''' ADDED BY SHWETA DATED 10.1.2012 FOR SKIPING GENERAL CASE TO CHECK DOCCODE
    If Trim(lbl_cat.Caption) <> "C" And Trim(lbl_cat.Caption) <> "NC" Then
    doc_code
    End If
      If RS_reqdet.RowCount = 0 Then
    MsgBox "No corresponding req. no. record exists or either its reqnstat is M in rrequisitn table"
  Else
    'me_rdate = rs_reqdet("reqdate") ' comm on 24/7
    ReDim Preserve UserData(0 To 5, 0 To 12)
    'DBGrid1.Refresh

    DBGrid1.col = 0
    'DBGrid1.row = 0
    If Trim(cbo_reqno.Text) <> "" Then
        DBGrid1.Text = cbo_reqno.Text & " " & RS_reqdet("reqndt") & " " & Trim(RS_reqdet("pat_catg"))
    End If
    'DBGrid1.Text = RS_reqdet("ser_code") & " " & reqno & " " & RS_reqdet("reqndt") & " " & Trim(RS_reqdet("pat_catg"))
    me_rdate.Text = RS_reqdet("reqndt")
    If flag = 1 Then
       lbl_cat = Trim(RS_reqdet("pat_catg"))
    End If
    DBGrid1.col = 4
    DBGrid1.Text = lbl_dcode
    DBGrid1.col = 5
    DBGrid1.Text = RS_reqdet("REFBYDOC")
    
    
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
          DBGrid1.SelStart = 24
End If
            DBGrid1.col = 0
            DBGrid1.Enabled = True
            DBGrid1.SetFocus
'ADD ON 24/7

If cbo_reqno.ListCount <> 0 Then

    If Trim(cbo_reqno.Text) <> "" Then
        cbo_reqno.RemoveItem cbo_reqno.ListIndex
    End If
'Else
 '  REQGOTFALG = False
  ' Me.Hide
End If
If cbo_reqno.Text = "" Then
'cbo_reqno.AddItem pathology.SHW_REQNO
End If
Exit Sub
Errhandler:
MsgBox Err.Description, vbCritical, "MEMO"
Exit Sub
End Sub

Private Sub Com_dno_GotFocus()
'Dim ind_cbodept As Integer
'If Trim(Com_dno) = "" Then
'    For ind_cbodept = 0 To (Com_dno.ListCount - 1)
'        If Trim(Com_dno.List(ind_cbodept)) = "RADIODIAGNOSIS$LZZ" Then
'            Com_dno.ListIndex = ind_cbodept
'        End If
'    Next ind_cbodept
'    End If
End Sub

Private Sub Com_dno_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Or KeyAscii = 9 Then
    Com_dno_LostFocus
End If
End Sub

Private Sub Com_dno_LostFocus()
On Error Resume Next
''''Set rs2 = rdb.OpenResultset("select dept_code from MA106 where dept_name = '" & UCase(Com_dno.Text) & "'", rdOpenKeyset)
''''lbl_dname = rs2("dept_code") '***31399
''''lbl_dname = Com_dno.ItemData(Com_dno.ListIndex)
'                rs12.MoveFirst
'                While Not rs12.EOF
'                    If rs12("DEPT_NAME") = Com_dno.Text Then
'                        lbl_dname = rs12("dept_CODE")
'                        rs12.MoveLast
'                    End If
'                    rs12.MoveNext
'                Wend
'
'    me_cno.SetFocus
End Sub

Private Sub commit_GotFocus()
VALIDATEAMT
End Sub

Private Sub d_u_name_Click()
txt_unit_doc_name.Text = d_u_name
Call doc_code

If rs2("pat_reg_ty") = "G" Then
    lbl_dcode = " "
ElseIf rs2("pat_reg_ty") = "P" Then
''' ADDED BY SANDHYA DATED 11.5.2011
''''' LISTING DOCTORS FROM PAYROLLLIB WHOSE CESSYN = 4
  SELMA102RS.Requery
  If SELMA102RS.RowCount > 0 Then
    While Not SELMA102RS.EOF
         d_u_name.AddItem Trim(SELMA102RS("DOC_NAME"))
         SELMA102RS.MoveNext
    Wend
  End If
End If

'NEW IMLPEMENTED FOR REFERAL PATIENT
If rs2("pat_reg_ty") = "R" Then
   If rs2("U_D") = "U" Then
         lbl_dcode = " "
   ElseIf rs2("U_D") = "D" Then
      ''''' ADDED BY SANDHYA DATED 11.5.2011
   ''''' LISTING DOCTORS WHOSE CESSYN = 4
    SELDOCCODE(0) = "%" & d_u_name.Text & "%"
    Set RSSELDCODE = SELDOCCODE.OpenResultset(2, 3)
    If RSSELDCODE.RowCount = 1 Then
       lbl_dcode = RSSELDCODE("DOC_CODE")
    End If
   
   End If
End If
cbo_reqno.SetFocus

'DBGrid1.COL = 0
'DBGrid1.SetFocus
MSG:
If Err = 0 Then Exit Sub
MsgBox "Doctor code not found"

Resume Next
End Sub



Private Sub DBGrid1_AfterColEdit(ByVal ColIndex As Integer)
On Error GoTo MSG
Dim opt As String
'dim ncchg as string 'commented 4899
If ColIndex = 0 Then
  DBgrid1_KeyPress (13)
  If ncchg = "Z" Then 'ADDED  foR z series memos 4899  'MODIFY THIS 23999
    DBGrid1.Columns(3).Locked = False
    DBGrid1.col = 2
   
  Else  'ADDED for z series memos 4899
    'DBGrid1.Columns(3).Locked = True
    
    DBGrid1.col = 2 'ADDED for z series 4899 'MODIFY THIS 23999
  End If 'ADDED for z series 4899
  
  
ElseIf ColIndex = 2 Then
        If rs18flag = False Then
            Ma101Qr(0) = DBGrid1.Columns(0)
            Set RS18 = Ma101Qr.OpenResultset(1, 3)
        Else
            Ma101Qr(0) = DBGrid1.Columns(0)
            RS18.MoreResults
            RS18.Requery
        End If
        ''''Set RS18 = rdb.OpenResultset("select * from MA101 where ser_code = '" & DBGrid1.Columns(0) & "'", rdOpenKeyset, 3)
        
        opt = Trim(lbl_cat)
        ncchg = Left(DBGrid1.Columns(0), 1) '***22/10/98
        
        Select Case opt
        Case "C"
            DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH")
            
        Case "B"
            If opttyp <> "AF" Then '***22/10/98
                DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("B_HOSP_CH"))
                
            Else
                DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("B_HOSP_CH") - RS18("B_PROF_CH")) '***22/10/98
                 
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
            
        Case "F"
            If opttyp <> "AF" Then '***22/10/98
                DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH"))
                 
            Else
                
                DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH") - RS18("FN_PROF_CH"))  '***22/10/98
                
            End If
        Case "NC"
            DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("NC_CH"))  '''''RRRR23899
             
        'If ncchg <> "I" Then '**22/10/98
        '    DBGrid1.Columns(3) = 0
        'Else                 '**22/10/98
        '    DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH") '**22/10/98
                'End If '**22/10/98
                
         ''''' ADDED BY SANDHYA DATED 6.07.2012 FOR BP PATIENT RGJAY
        Case "BP"
            DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH")
            
        Case Else
            MsgBox "NO MATCH ..."
        End Select
        
         If ncchg = "Z" Then 'ADDED 27799
          DBGrid1.col = 3
        Else
          DBGrid1.col = 4
        End If   'ADDED 4899
        DBGrid1.SetFocus
     
ElseIf ColIndex = 4 Then
    
    '''Set rs_dname = rdb.OpenResultset("select * from ma102 where doc_code = " + DBGrid1.Columns(4).Text, 2, 3)
  
    If rs_dname.RowCount <> 0 Then
        DBGrid1.col = 5
        '''DBGrid1.Text = RS_DNAME("doc_name")
        
        rs_dname.MoveFirst
        While Not rs_dname.EOF
            If rs_dname("doc_code") = DBGrid1.Columns(4).Text Then
                DBGrid1.Text = rs_dname("doc_name")
                rs_dname.MoveLast
            End If
            rs_dname.MoveNext
        Wend

        DBGrid1.row = DBGrid1.row + 1
        DBGrid1.col = 0
    Else
       MsgBox "NO CORRESPONDING DOCTOR NAME EXISTS....."
       DBGrid1.col = 4
       DBGrid1.SelStart = 0
       DBGrid1.SelLength = 5
       DBGrid1.SetFocus
    End If
Else
Exit Sub
End If
Exit Sub

MSG:
DBGrid1.col = 0

DBGrid1.SetFocus
Exit Sub
Text22.Text = " PRESS CONTROL KEY TO COME OUT OF DBGRID .... and TO REFRESH TOTAL AMOUNT" '****9499
End Sub

Private Sub DBGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, cancel As Integer)

ncchg = Left(DBGrid1.Columns(0), 1) ' added for z-series memos***4899
If ColIndex = 3 Then
    If ncchg = "Z" Then
        Exit Sub
    Else ' added for z-series memos***4899
        MsgBox "YOU CANNOT MODIFY AMOUNT FOR " & CStr(DBGrid1.Columns(0)) & "   SERVICE"  ' added for z-series memos***4899
        'KeyAscii = 0
        cancel = True
        Exit Sub
    End If
    Exit Sub
End If ' added for z-series memos***4899

End Sub

Private Sub DBGrid1_Click()
DBGrid1.Columns(1).Locked = True
'DBGrid1.Columns(3).Locked = True ' commented for z-series memos***4899

End Sub

Private Sub DBGrid1_GotFocus()
On Error GoTo MSG:
'DBGrid1.Columns(5).Locked = True
Text22 = " PRESS CONTROL KEY TO COME OUT OF DBGRID ...."
'If DateDiff("d", Format(rs2("consult_dt"), "DD/MM/YYYY"), me_date) > 90 Then
     
 '    I = MsgBox(DateDiff("d", Format(rs2("consult_dt")), me_date), , "NO OF DAYS : ")
  '   MsgBox "SELECT 'B003' CODE, FOR RECONSULTATION CHARGES...."
'End If
DBGrid1.col = 0
DBGrid1.SetFocus
Exit Sub
MSG:
End Sub


Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 17 Then
    If txt_rem.Enabled = True Then ''5599
        txt_rem.SetFocus '5599
    ElseIf New1.Visible = True Then ''5599
        New1.SetFocus '5599
    End If '5599
End If
End Sub


Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo MSG
If DBGrid1.col = 5 And KeyCode = 13 Then
'   MsgBox Len(Trim(DBGrid1.Text))
   If Len(Trim(DBGrid1.Text)) = 0 Then
        If rs18flag = False Then
            Ma101Qr(0) = DBGrid1.Columns(0)
            Set RS18 = Ma101Qr.OpenResultset(1, 3)
            rs18flag = True
        Else
            Ma101Qr(0) = DBGrid1.Columns(0)
            RS18.MoreResults
            RS18.Requery
        End If
        ''''Set RS18 = rdb.OpenResultset("select * from MA101 where ser_code = '" & DBGrid1.Columns(0) & "'", rdOpenKeyset, 3)
        opt = lbl_cat

        Select Case opt
        Case "B "
                If RS18("B_PROF_CH") <> 0 Then
                    MsgBox "ENTER DOCTOR CODE....."
                    DBGrid1.col = 4
                    DBGrid1.SelLength = 5
                    DBGrid1.SetFocus
                Else
                    DBGrid1.row = DBGrid1.row + 1
                    DBGrid1.col = 0
                End If
  
        Case "A "
           If RS18("A_PROF_CH") <> 0 Then
              MsgBox "ENTER DOCTOR CODE....."
              DBGrid1.col = 4
              DBGrid1.SelLength = 5
              DBGrid1.SetFocus
           Else
               DBGrid1.row = DBGrid1.row + 1
               DBGrid1.col = 0
           End If
            
        Case "D "
           If RS18("D_PROF_CH") <> 0 Then
              MsgBox "ENTER DOCTOR CODE....."
              DBGrid1.col = 4
              DBGrid1.SelLength = 5
              DBGrid1.SetFocus
           Else
               DBGrid1.row = DBGrid1.row + 1
               DBGrid1.col = 0
           End If
        Case "F "
           If RS18("FN_PROF_CH") <> 0 Then
              MsgBox "ENTER DOCTOR CODE....."
              DBGrid1.col = 4
              DBGrid1.SelLength = 5
              DBGrid1.SetFocus
           Else
               DBGrid1.row = DBGrid1.row + 1
               DBGrid1.col = 0
           End If
        
        Case Else
            DBGrid1.row = DBGrid1.row + 1
            DBGrid1.col = 0
            End Select
    End If
End If
Exit Sub
MSG:
DBGrid1.col = 0
DBGrid1.SetFocus
Exit Sub

End Sub

Private Sub DBGrid1_LostFocus()
Dim tot As Long
Dim iRow As Integer
Text22 = " "
tot = 0
For iRow = 0 To mTotalRows&
'For iRow = 0 To mTotalRows& - 1 ''''''''should be removed
If Len(Trim(UserData(0, iRow))) > 0 Then '''''25041999
    If Not IsEmpty(UserData(0, iRow)) Then
    ''' ADDED SANDHYA PROTOCOL PATIENT
     If PROTFLG = True Then
        prottot = CDbl(PTOT)
     End If
           tot = tot + CLng(UserData(3, iRow))
     End If
     '' END OF PROTOCOL
End If
Next iRow
'''lbl_tot_amt = Format(tot, "#,##,##0.00")

lbl_tot_amt = CDbl(tot)
''' SANDHYA PROTOCOL PATIENT
If PROTFLG = True Then
nettot = lbl_tot_amt - prottot
ElseIf PTOT > 0 And NPROTFLG = True Then
prottot = CDbl(PTOT)
nettot = lbl_tot_amt - PTOT
Else
nettot = lbl_tot_amt
End If
'' end of protocol
End Sub



Private Sub delete_Click()
'If Trim(USRGRP) <> "GROUP3" Then
' MsgBox "ONLY USERS OF GROUP3 ARE ALLOWED TO DELETE, YOUR GROUP IS : " & USRGRP, vbOKOnly
' Exit Sub
'End If

MsgBox "Rights For This Option Are Not Available"
flag1 = 3
'view_Click
'del
End Sub

Private Sub exit1_Click()
'Call exit_prog
Unload Me
End
End Sub

Private Sub DBgrid1_KeyPress(KeyAscii As Integer)
On Error Resume Next 'GoTo msg
Dim rs_dcode As rdoResultset
Dim ncchg As String ' commented for z-series memos***4899

Dim opt As String
reflg = 0 '***24/11/98
opttyp = lbl_typ '***22/10/98
ncchg = Left(DBGrid1.Columns(0), 1) '***22/10/98
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If KeyAscii = 13 Then
    If DBGrid1.col = 3 Or DBGrid1.col = 4 Or DBGrid1.col = 5 Then  ' added for z-series memos***4899

''' '===========================================================
' If DBGrid1.Col = 4 Then
'Set RS18 = rdb.OpenResultset("select * from MA101 where ser_code = '" & DBGrid1.Columns(0) & "'", rdOpenKeyset, 3)
'opt = lbl_cat
'ncchg = Left(DBGrid1.Columns(0), 1) '***22/10/98
'
'Select Case opt
'
'Case "C "
'
'Case "B "
'
'
'    If RS18("B_PROF_CH") <> 0 And DBGrid1.Columns(LastCol) = "" Then
'       MsgBox "ENTER THE DOCTOR CODE"
'       DBGrid1.row = LastRow
'       DBGrid1.Col = LastCol
''      DBGrid1.Col = DBGrid1.Col + 1
'       DBGrid1.SetFocus
'      End If
'
'
'
'Case "A "
'    If RS18("A_PROF_CH") <> 0 And DBGrid1.Columns(LastCol) = "" Then
'       MsgBox "ENTER THE DOCTOR CODE"
'       DBGrid1.row = LastRow
'       DBGrid1.Col = LastCol
''      DBGrid1.Col = DBGrid1.Col + 1
'       DBGrid1.SetFocus
'      End If
'
'
'Case "D "
'    If RS18("D_PROF_CH") <> 0 And DBGrid1.Columns(LastCol) = "" Then
'       MsgBox "ENTER THE DOCTOR CODE"
'       DBGrid1.row = LastRow
'       DBGrid1.Col = LastCol
''      DBGrid1.Col = DBGrid1.Col + 1
'       DBGrid1.SetFocus
'      End If
'
'
'Case "F "
'    If RS18("FN_PROF_CH") <> 0 And DBGrid1.Columns(LastCol) = "" Then
'       MsgBox "ENTER THE DOCTOR CODE"
'       DBGrid1.row = LastRow
'       DBGrid1.Col = LastCol
''      DBGrid1.Col = DBGrid1.Col + 1
'       DBGrid1.SetFocus
'      End If
'
'Case "NC"
'
'End Select
'Exit Sub
'End If
'
''       IF DBGRID1.Columns(LASTCOL).TEXT
'    '==========================================================

        If DBGrid1.col = 5 Then
            DBGrid1.row = DBGrid1.row + 1
            DBGrid1.col = 0
            DBGrid1.SetFocus
        End If
            Exit Sub
        End If ' added for z-series memos***4899
        If rs18flag = False Then
            Ma101Qr(0) = DBGrid1.Columns(0)
            Set RS18 = Ma101Qr.OpenResultset(1, 3)
            rs18flag = True
        Else
            Ma101Qr(0) = DBGrid1.Columns(0)
            RS18.MoreResults
            RS18.Requery
        End If
        ''''Set RS18 = rdb.OpenResultset("select * from MA101 where ser_code = '" & DBGrid1.Columns(0) & "'", rdOpenKeyset, 3)
        If RS18.RowCount = 0 Then
            MsgBox " SERVICE CODE DOES NOT EXIST IN THE MASTER TABLE ...."
            DBGrid1.col = 0 'ADDED 23999
            DBGrid1.SelStart = 5 '23899
            DBGrid1.SetFocus
            Exit Sub '23899
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
            
            opt = Trim(lbl_cat)
        ''' sandhya protocol patient starts
        
        PROTCASE_QR(0) = CASENO
        PROTCASE_QR(1) = MIRBNO
        Set PROTCASE_RS = PROTCASE_QR.OpenResultset(2, 3)
        PROTCASE_RS.Requery
        If PROTCASE_RS.RowCount > 0 Then
           PROTSERV_QR(0) = MIRBNO
           'PROTSERV_QR(0) = Trim(PROTCASE_RS("IRBno"))
           PROTSERV_QR(1) = Left(DBGrid1.Columns(0), 4)
           Set ZAS400DTTM = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1,CURRENT TIME TM1 FROM PATOOLSLIB.MA107") 'added for date
          ' ZAS400DTTM.Requery
'           PROTSERV_QR(2) = Format(ZAS400DTTM!DT1, "YYYY-MM-DD")
'           PROTSERV_QR(3) = Format(ZAS400DTTM!DT1, "YYYY-MM-DD")
           Set PROTSERV_RS = PROTSERV_QR.OpenResultset(2, 3)
           PROTSERV_RS.Requery
          If PROTSERV_RS.RowCount = 1 Then
            If Trim(PROTSERV_RS("FLD1")) = "E" Then
              DBGrid1.Columns(3) = PROTSERV_RS("ser_amt")
             ' PTOT = 0
              PTOT = PTOT + DBGrid1.Columns(3)
              PROTFLG = True
              '***** 09/02/10 *****
               Set RSPINAME = rdb.OpenResultset("SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE IN (SELECT PICCNO FROM IRBLIB.PRJDTLS WHERE PERPRJNO = '" & MIRBNO & "')", 2, 1)
               
               If RSPINAME.RowCount = 1 Then
                    PINAME = Trim(RSPINAME("FNAME")) & " " & Trim(RSPINAME("MNAME")) & " " & Trim(RSPINAME("LNAME"))
                Else
                   PINAME = ""
               End If
              '********************
            ElseIf Trim(PROTSERV_RS("FLD1")) = "I" Then
              PROTFLG = True
             
             Ma101Qr(0) = DBGrid1.Columns(0)
             Set RS18 = Ma101Qr.OpenResultset(1, 3)
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
            
            Case "F" ''5599
                If opttyp <> "AF" Then '***22/10/98
                    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH"))
                
                Else
                    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH") - RS18("FN_PROF_CH"))  '***22/10/98
               
                End If
            Case "NC"
        
                DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("NC_CH"))  '''''RRRR23899
            Case "BP"
                DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH")
            
            Case Else
                MsgBox "NO MATCH ..."
            End Select
            End If
         Else
            NPROTFLG = True
            Ma101Qr(0) = DBGrid1.Columns(0)
             Set RS18 = Ma101Qr.OpenResultset(1, 3)
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
            
            Case "F" ''5599
                If opttyp <> "AF" Then '***22/10/98
                    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH"))
                
                Else
                    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH") - RS18("FN_PROF_CH"))  '***22/10/98
               
                End If
            Case "NC"
        
                DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("NC_CH"))  '''''RRRR23899
                
            ''''' ADDED BY SANDHYA DATED 6.07.2012 FOR BP PATIENT RAJIV GANDHI
            Case "BP"
            DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH")
            
            Case Else
                MsgBox "NO MATCH ..."
            End Select
                                   
       ' End If
     End If
    Else
        NPROTFLG = True
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

        Case "F" ''5599
            If opttyp <> "AF" Then '***22/10/98
                DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH"))

            Else
                DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH") - RS18("FN_PROF_CH"))  '***22/10/98

            End If
        Case "NC"

            DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("NC_CH"))  '''''RRRR23899
        Case "BP"
            DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH")

        '    If ncchg <> "I" Then '**22/10/98
        '        DBGrid1.Columns(3) = 0
        '    Else                 '**22/10/98
        '        DBGrid1.Columns(3) = DBGrid1.Columns(2) * RS18("C_CH") '**22/10/98
        '    End If '**22/10/98

        Case Else
            MsgBox "NO MATCH ..."
        End Select
       End If
      End If
        DBGrid1.Columns(4).Locked = True
        DBGrid1.Columns(4) = lbl_dcode
        DBGrid1.Columns(5) = txt_unit_doc_name
        DBGrid1.col = DBGrid1.col + 1
        DBGrid1.SetFocus
     
        
        'DBGrid1.Row = DBGrid1.Row + 1
   ' End If
    'Exit Sub ' commented on 9499
End If
'msg: 'MsgBox " SERVICE CODE DOES NOT EXIST IN THE MASTER TABLE ...."
'DBGrid1.SetFocus
End Sub

Private Sub flex2_Click()
On Error GoTo MSG
Dim opt As String

'''Set rs5 = rdb.OpenResultset("select * from MA102 where doc_code = " & flex2.Text, 2, 3)  ' & " or doc_code = " & Val(DBGrid1.Columns(4).Text), rdOpenKeyset, 3) ' 9/6/99 COMMENTED FROM OR DOCCODE

If rs_dname.RowCount <> 0 Then
    rs_dname.MoveFirst
    While Not rs_dname.EOF
        If rs_dname("doc_code") = flex2.Text Then
            txt_unit_doc_name = rs_dname("UD_NAME")
            DBGrid1.Columns(5) = rs_dname("doc_name")
            DBGrid1.Columns(4) = rs_dname("doc_code")
            lbl_dcode = rs_dname("doc_code")
            rs_dname.MoveLast
        End If
        rs_dname.MoveNext
    Wend
Else
    Exit Sub
End If  '9/6/99 ADDED
flex2.Visible = False
DBGrid1.row = DBGrid1.row + 1
'FLEX2FLAG = True
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

Private Sub Form_Unload(cancel As Integer) '''A251199
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

Private Sub M_number_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0 'M_number1_LostFocus
End If
End Sub

Private Sub M_number_LostFocus()
On Error GoTo MSG
''Dim st As String
''Dim I As Integer
''Dim row
''If flag = 2 Then
''ReDim UserData(0 To 5, 0 To 12)
''DBGrid1.Refresh
''End If
''DT = Mid(M_number, 2, 1)
''    chargen = Left(Me.M_number.Text, 2) 'ADDED 11/11/2001
''    'M_number.Text = padzero(6, "OV", Mid(M_number, 3), 1)
''M_number.Text = padzero(6, "M", Mid(M_number, 3), 1)
''If M_number.Text <> " " Then
''    ''''st = "select * from TA1071 where memo_no = '" & "M-" & M_number.Text & "'"
''    Ta1071Qr(0) = M_number.Text '"M-" & M_number.Text
''    Set rs7 = Ta1071Qr.OpenResultset(2, 3)
''    ''''Set rs7 = rdb.OpenResultset(st, rdOpenKeyset, 3)
''    If rs7.RowCount = 0 Then
''        MsgBox "RECORD NOT AVAILABLE.", vbCritical
''        Exit Sub
''    End If
''
''    'com_mno.Locked = True
''    Ta107Qr(0) = M_number.Text  '"M-" & M_number.Text
''    Ta107Qr(1) = "O"
''    Set rs10 = Ta107Qr.OpenResultset(2, 3)
''    ''''Set rs10 = rdb.OpenResultset("select * from TA107 where memo_no = '" & "M-" & M_number.Text & "'", rdOpenKeyset)
''
''     If rs10.RowCount Then
''        If Trim(rs10("DIS_FLAG")) = "D" Then
''            MsgBox "DISCOUNT IS GIVEN FOR THIS MEMO"
''            btn_vis (0)
''            cleartxt
''            Exit Sub
''        End If
'''      '**** Inpatient checking
'''       If rs10("IO_MEMO") = "I" Then
'''            MsgBox "Inpatient Memo's cannot be viewed in Out patient Memo's"
'''            btn_vis (0)
'''            cleartxt
'''            Exit Sub
'''        End If
'''     '**** Inpatient checking
''        me_date.Text = Format(rs10("memo_date"), "dd/mm/yyyy")
''        'lbl_dno = rs10("dept_code")
''        lbl_dname = rs10("dept_code")
''        lbl_cno = rs10("case_no")
''        lbl_cat.Caption = rs10("MEMO_catg")  '9/6/99 ADDED BECAUSE TO VIEW IP/OP MEMOS PATIENT CATG SHOULD BE THE CATEGORY AT THE TIME OF MEMO RAISING WHERE AS TA101 CATG IS LATEST CATG
''     Else
''           MsgBox "Inpatient Memo's cannot be viewed in Out patient Memo's"
''            btn_vis (0)
''            cleartxt
''            Exit Sub
''     End If
''
''    lbl_dcode.Enabled = False
''    txt_unit_doc_name.Enabled = False
''
''    DBGrid1.Columns(0).Locked = True
''    DBGrid1.Columns(2).Locked = True
''    d_u_name.Visible = False
''    txt_doc_sec_name.Enabled = False
''    txt_rem.Enabled = False
''    Call patient_det(lbl_cno)
''        If Not IsNull(rs10("doc_code")) Then
''            lbl_dcode = rs10("doc_code")
''            If rs2("pat_reg_ty") = "P" Then
''                    '''Set rs1 = rdb.OpenResultset("select doc_name from MA102 where doc_code = " + lbl_dcode, rdOpenKeyset)
''               txt_unit_doc_name = rs_dname("doc_name")
''            End If
''         Else
''                MsgBox (" Referred to unit ")
''         End If
''    txt_rem = rs10("rem")
''     '-------------------------------------------------------
''     '06/04/2002-Parag
''     'For Displaying the receipt details if payment is made
''     'against the memo.
''      zmemorcpt(0) = M_number.Text
''      Set zrsmemorcpt = zmemorcpt.OpenResultset(2, 1)
''      If zrsmemorcpt.RowCount > 0 Then
''         txt_rem.Text = zrsmemorcpt!REC_NO + " On: " + CStr(zrsmemorcpt!rec_date)
''         txt_rem.Refresh
''         zrsmemorcpt.Close
''         Set zrsmemorcpt = Nothing
''      End If
''     '-------------------------------------------------------
''
''    lbl_tot_amt = Format(rs10("memo_amt"), "#,##,##0.00")
''        If Not IsNull(rs10("d_s_name")) Then
''            txt_doc_sec_name = rs10("d_s_name")
''        End If
''        If Not IsNull(lbl_dname) Then
''                '''Set rs2 = rdb.OpenResultset("select dept_name from MA106 where dept_code = '" & UCase(lbl_dname) & "'", rdOpenKeyset)
''                rs12.MoveFirst
''                While Not rs12.EOF
''                    If rs12("DEPT_CODE") = UCase(lbl_dname) Then
''                        lbl_dno = rs12("dept_name")
''                        rs12.MoveLast
''                    End If
''                    rs12.MoveNext
''                Wend
''                '''lbl_dno = rs2("dept_name")
''        End If
''    If rs7.RowCount Then
''        rs7.MoveFirst
''        For row = 0 To rs7.RowCount - 1
''        If Len(Trim(rs7(6).Value)) = 0 Then
''        UserData(0, row) = rs7(1).Value
''        Else
''        UserData(0, row) = rs7(1).Value & "/" & rs7(6).Value & "/" & rs7(7).Value & "/" & rs7(8).Value
''       End If
''        If Not IsNull(rs7(1).Value) Then
''
''            Ma101Qr(0) = rs7(1).Value
''            Set rs13 = Ma101Qr.OpenResultset(2, 3)
''            'MsgBox rs13.RowCount
''            ''Set rs13 = rdb.OpenResultset("select ser_det from MA101 where ser_code = '" & rs7(1).Value & "'", rdOpenKeyset, 3)
''
''            UserData(1, row) = rs13("ser_det")
''            UserData(2, row) = CStr(rs7(3).Value)
''            UserData(3, row) = CStr(rs7(2).Value)
''                If Not rs7(4).Value = 0 Then
''                    UserData(4, row) = CStr(rs7(4).Value)
''                    '''Set rs_dname = rdb.OpenResultset("select doc_name from ma102 where doc_code = " + CStr(rs7(4).Value), 2, 3)
''                    '''UserData(5, row) = RS_DNAME("doc_name")
''                        If rs2("pat_reg_ty") = "P" Then
''                            rs_dname.MoveFirst
''                            'While Not rs_dname.EOF
''                                'If rs_dname("doc_code") = CStr(rs7(4).Value) Then
''                                    UserData(5, row) = rs_dname("doc_name")
''                                    'rs_dname.MoveLast
''                                'End If
''                                'rs_dname.MoveNext
''                            'Wend
''                        End If
''                End If
''                If Len(Trim(rs7(6).Value)) = 0 Then
''                txt_req_no = ""
''                me_rdate.Text = ""
''                Else
''                txt_req_no = rs7(6).Value
''                me_rdate.Text = rs7(7).Value
''                End If
''        End If
''        rs7.MoveNext
''        Next row
''    End If
''    DBGrid1.Refresh
''End If
''    If flag1 = 3 Then
''    'Text22 = "press TAB to get the delete message...."
''    del
''    flag1 = 0
''    End If
 Exit Sub
''
MSG: MsgBox "ENTER VALID MEMO NUMBER "
MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub me_cno_GotFocus()
'If FLEX2FLAG = True Then
'      DBGrid1.SetFocus
'      Exit Sub
' End If
'

If Len(Trim(Com_dno.Text)) = 0 Then
    Com_dno.Enabled = True
    Com_dno.SetFocus
End If
End Sub

Private Sub me_cno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'Or KeyAscii = 9 Then
    KeyAscii = 0 'me_cno_LostFocus
End If
End Sub

Private Sub me_cno_LostFocus()
'''MsgBox Err
''Dim RSINPT As rdoResultset
'On Error GoTo MSG
''Dim c As Integer
''Dim i
'''If FLEX2FLAG = True Then
'''      DBGrid1.SetFocus
'''      Exit Sub
''' End If
''If flag = 1 Then
''    If Len(Trim(me_cno.Text)) <> 0 Then
''        NoCaseFlag = 0
''        me_cno = pad(5, "C", me_cno.Text, 3)
''        CASENO = Left(me_cno, 2) & "/" & Right(me_cno, 5)
''        '''Set RS15 = rdb.OpenResultset("select count(*) as c  from ta101 where case_no = '" & CASENO & "'", rdOpenKeyset, 3) '''C251199
''        OPQ(0) = CASENO '''A251199
''        rs2.MoreResults
''        rs2.Requery
''        ''''Set rs2 = OPQ.OpenResultset(rdOpenKeyset)
''        If rs2.RowCount = 0 Then  '''M251199
''                NoCaseFlag = 1
''                i = MsgBox("CASE NUMBER DOES NOT EXIST, DO YOU WANT TO CONTINUE ", vbYesNo)
''                If i = vbYes Then
''                   me_cno.SelStart = 0  '***020499
''                   me_cno.SelLength = 8
''                   me_cno.SetFocus   '*****020499
''                Else
''                   CANCEL_Click
''                   Exit Sub
''                End If
''        Else
''             'Set RSINPT = rdb.OpenResultset("select count(*) as c  from ta105 where case_no = '" & CASENO & "' AND DISC_DATE IS NULL", rdOpenKeyset, 3)'**COMM ON 26/8/99
''             InPt_ChkQr(0) = CASENO
''             Set RSINPT = InPt_ChkQr.OpenResultset(2, 3)
''             ''''Set RSINPT = rdb.OpenResultset("select count(*)  from ta105 where case_no = '" & CASENO & "' AND DISC_DATE IS NULL AND ADMS_DATE IS NOT NULL", rdOpenKeyset, 3) '**ADDED ON 26/8/99
''             If RSINPT.RowCount > 0 Then   'ADDED 23899
''             ''' CHECKING FOR INPATIENTS OF CASH PAYING & NO CHARGE
''             '23/01/2002-Parag
''             'For "NC" patients irrestpective of whether they are OUT-PATIENTS/IN-PATIENTS
''             'and CP/TA/CR only Outpatient memos are to be generated.
''              If (rs2("PAT_TYPE") = "CP" Or rs2("PAT_TYPE") = "TA" Or rs2("PAT_TYPE") = "CR") And rs2("PAT_CATG") = "NC" Then
'''               PATSTR = "CASE NO. " + CASENO + " IS  CASH PAYING  AND NC TYPE DO YOU WANT TO GENERATE MEMOS ?"
''                MsgBox "CREATING MEMO FOR IN-PATIENT OF TYPE ====> " & rs2("PAT_TYPE") & " |  NC"
''                 INNC = True
'''               If RESP = vbNo Then
'''
''              Else
''                'txt_unit_doc_name.Enabled = False
''                MsgBox "OUTPATIENT MEMOS CANNOT BE GIVEN TO INPATIENTS"
''                CANCEL_Click
''                cleartxt
''                'txt_unit_doc_name.Enabled = True
''                Exit Sub
''              End If
''
''             End If  'ADDED 23899
''
''
'On Error GoTo PATERR:
''
''         Call patient_det(CASENO) '1
''
''If flag = 1 Then
''    txt_req_no.ZOrder 1
'''    RQ_REQDETREQNO(0) = CASENO
'''    RQ_REQDETREQNO(1) = "OP"
'''   RQ_REQDETREQNO(2) = lbl_dname.Caption
''
''' remove the comments here and made the comments above chandu
''
''    '23/01/2002-Parag.
''    'For NC INPATIENTS in RISLIB/RREQUISITN the
''    'Since Only OP memos are to be generated for NC patients
''    'The Parameter passed is "IP" if INNC flag is True For "NC" patients
''    lbl_typ.Refresh
''    If Trim(lbl_cat.Caption) <> "NC" And INNC = True Then
''    RQ_REQDETREQNO(0) = CASENO
''    RQ_REQDETREQNO(1) = "IP"
''    INNC = False
''    ElseIf Trim(lbl_cat.Caption) = "NC" And INNC = True Then
''            RQ_REQDETREQNO(0) = CASENO
''            RQ_REQDETREQNO(1) = "IP"
''        Else
''            RQ_REQDETREQNO(0) = CASENO
''            RQ_REQDETREQNO(1) = "OP"
''    End If
''
''
''   Set RS_REQDETREQNO = RQ_REQDETREQNO.OpenResultset(1, 1)
''    'MsgBox RS_REQDETREQNO.RowCount
'' cbo_reqno.clear
''While Not RS_REQDETREQNO.EOF
'' cbo_reqno.AddItem req_chglen(RS_REQDETREQNO("reqnno"))
'' RS_REQDETREQNO.MoveNext
''Wend
''cbo_reqno.Enabled = True
''cbo_reqno.SetFocus
''End If
''
''
'On Error GoTo UDERR:
''
''         Call ud_name
''      End If
''   End If
''Else
'' cancel.SetFocus
''End If
'''If (CASENO = "ZA/99999") Or (CASENO = "ZB/99999") Or (CASENO = "ZD/99999") Or (CASENO = "ZF/99999") Then
''' If lbl_cat = "C" Or lbl_cat = "NC" Then
'''  'NO OPERATION TO BE DONE
'''  ' OTHERWISE GET DOC NAME
'''
''' Else
'''       op_memo_d1.Hide
'''       BLOODCOMBO.ZOrder
'''        DoEvents
'''       BLOODCOMBO.Left = 500
'''       BLOODCOMBO.Top = 2000
'''       BLOODCOMBO.Show vbModal
'''
'''       DoEvents
'''       BLOODCOMBO.BLOODDOC.SetFocus
'''       txt_unit_doc_name.Text = BLOODCOMBO.BLOODDOC.Text
'''       op_memo_d1.Show
''' End If
'''
'''End If
''
''
''
''
''
'''
'''
''' End If
'''MsgBox "HERE IN at me_cno lostfocus"
''d_u_name.Text = txt_unit_doc_name
''If Trim(d_u_name) <> "" Then
'On Error GoTo DUERR:
''
''d_u_name_Click
''
''
''
''''' CHECKING FOR PATIENT CASENOS ZA/99999, ZB/99999,ZC/99999,ZD/99999,ZF/99999
''''''' GIVE MSGBOX TO ENTER ONLY BLOOD SALE RELARED SER_CODES
'''If (CASENO = "ZA/99999") Or (CASENO = "ZB/99999") Or (CASENO = "ZD/99999") Or (CASENO = "ZF/99999") Then
'''    MsgBox " THIS IS OUTSIDE BLOOD BANK MEMOS, CHOOSE THE CORRECT DOCTOR NAME FROM THE COMBO"
'''
'''
'''End If
''End If
''If Trim(rs2("pat_reg_ty")) = "P" Then
''
''d_u_name.SetFocus
''Else
''cbo_reqno.SetFocus
''End If
''
'''cbo_reqno.SetFocus
''Exit Sub
''
''PATERR:
''        MsgBox "PATERR"
''        Exit Sub
''UDERR:
''        MsgBox "UDERR"
''        Exit Sub
''DUERR:
''
''      MsgBox "DUERR"
''        Exit Sub
''msg:
''  'If Err <> 0 Then
''  'MsgBox "ERROR IN CASE NUMBER"
''  'End If
End Sub

Private Sub me_mno_GotFocus()
Text22 = "PRESS ENTER KEY ....."
End Sub

Private Sub me_mno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ' Or KeyAscii = 9 Then
    KeyAscii = 0 'me_mno_1_LostFocus
End If
End Sub

Private Sub me_mno_LostFocus() ' ***instead of lostfocus event use separate private sub ***9499
Dim c As Integer
Text22 = " "
'K = Mid(me_mno.Text, 3, 8) ' ADDED 11/11/2001
'k1 = K
'If Len(Trim(Mid(me_mno.Text, 3))) <> 0 Then
''    me_mno = pad(6, "M", me_mno, 1)
'
' If RS_SERIES.RowCount <> 0 Then   ' ADDED 11/11/2001
'    me_mno = padzero(6, "M", K, 1)  ' ADDED 11/11/2001
'  Else                              ' ADDED 11/11/2001
'      me_mno = padzero(6, "M", K, 1)
'  End If
'     Exist_ChkQr(0) = me_mno.FormattedText
'   ' Exist_ChkQr(0) = "M-" & me_mno
'    Set rs22 = Exist_ChkQr.OpenResultset(2, 3)
'    ''''Set rs22 = rdb.OpenResultset("select count(*) as c  from ta107 where memo_no = '" & "M-" & me_mno & "'", rdOpenKeyset)
'    ''''If rs22("c") = 0 Then
'    If rs22.RowCount = 0 Then
'       Com_dno.Visible = True
'       Com_dno.SetFocus
'    Else
'        MsgBox "MEMO NUMBER ALREADY EXISTS ...... "
'        CANCELFLG = MsgBox("Dou you want to cancel ?", vbYesNo) '****31399
'        If CANCELFLG = vbYes Then
'            CANCEL_Click        '*******
'            Exit Sub '****31399
'        End If '****31399
'
'        me_mno.SelStart = 2 '***16/10/98
'        me_mno.SelLength = 8
'        me_mno.SetFocus '***16/10/98
'
'     End If
'Else
' CANCEL_Click
'
'End If
End Sub

Private Sub new1_Click()
'MsgBox Err
REQGOTFALG = False 'Added by Shraddha on 29/10/09 as told by Mrs.SRJoshi
'On Error GoTo MSG
EmptyGrid = 0 ' ***3499
AmtNotEq = 0 '***27599
flag1 = 0
flag = 1
'FLEX2FLAG = False
commit.Enabled = True
cancel.Enabled = True
'DBGrid1.Enabled = True
DBGrid1.Columns(0).Locked = False
DBGrid1.Columns(2).Locked = False
d_u_name.Visible = True
me_mno.Visible = True
M_number.Visible = False '***01399
me_cno.Visible = True
lbl_cno.Visible = False
Com_dno.Visible = True
'***** Added on 12/03/09 by Shraddha Pandit, as told by Sandhya M'am *****
PATDEPO = False
'**************************************************************************
'''****5599
' txt_unit_doc_name.Enabled = True
' txt_doc_sec_name.Enabled = True
' Com_dno.Enabled = True
' d_u_name.Enabled = True
'''****5599
'me_cno.Enabled = True
'txt_doc_sec_name.Enabled = True
'txt_rem.Enabled = True
'
'me_date.Enabled = True
'txt_unit_doc_name.Enabled = True
Enblbox (1) ''5599
cleartxt


me_mno.ZOrder
' COMMENTED FOR AUTOGENERATION OF MEMO NUMBER
'me_mno.SetFocus
'ME_mno.Mask = "M-######"
' COMMENTED FOR AUTOGENERATION OF MEMO NUMBER
me_cno.Mask = ">??/#####"
me_cno.Text = "__/_____"
me_cno.ZOrder
ReDim UserData(0 To 5, 0 To 12)
DBGrid1.Refresh
DBGrid1.Columns(0).Button = True 'ADD ON 24/7
btn_vis (1)
me_date = Format(DATE11, "dd/mm/yyyy") 'Changed from date to DATE11 by shraddha on 01/01/10 as told by Mrs.Sandhya

'**** AUTO GENERATION OF MEMO NUMBERS

'me_mno.Mask = ">?#######"
's = Right(me_date.Text, 4)
'RQ_SERIES(0) = "A" & s
'RS_SERIES.Requery
'
'If RS_SERIES.RowCount > 0 Then
'me_mno.Text = "M" & Right(RS_SERIES("YEAR1"), 1) & (Trim(RS_SERIES("SERIES")) + 1)
'
'' If IsNull(RS_SERIES("MEMO_SRNO")) Then
''     me_mno = RS_SERIES("CHAR_COR_YEAR") & "" & "1"
''    Else
''   me_mno.Text = RS_SERIES("CHAR_COR_YEAR") & "O" & (RS_SERIES("MEMO_SRNO") + 1)
'' End If
'Else
'   me_mno.Text = "M" & Right(me_date.Text, 1) & "1"
'End If
'

'**** AUTO GENERATION OF MEMO NUMBERS
'Set rs8 = rdb.OpenResultset("select max(memo_no) as v_mno from TA107", rdOpenKeyset)
'Set rs4 = rdb.OpenResultset("select * from TA107 where io_memo = 'O'", rdOpenKeyset, 3)
'''Set rs4 = rdb.OpenResultset("select * from TA107 where 1 = 2", rdOpenKeyset, 3) '''C221199
'a = Right(tempbuf3, 6)
'a1 = a + 1
'me_mno.Text = "M-" & a1
btn_vis (1)
me_date = Format(Date, "dd/mm/yyyy")
'me_mno.SelStart = 2
'me_mno.SetFocus
 'rs12.MoveFirst
me2 = rs12("dept_name")
 Com_dno.Text = me2
 Com_dno.SetFocus
 txt_doc_sec_name = usr_name
 'Exit Sub
'msg:   MsgBox Error(Err)
End Sub
Private Sub CANCEL_Click()
On Error GoTo MSG:
'   If flag = 1 Then
'        'rs4.AddNew
'        'rs4.CancelUpdate
'        'rs7.AddNew
'        'rs7.CancelUpdate
'   Else
'        rs4.Edit
'        rs4.CancelUpdate
'        rs7.Edit
'        rs7.CancelUpdate
'   End If
   cleartxt
   btn_vis (0)
   Enblbox (0) '''5599
lbl_smartcard.Visible = False
lbl_deposit.Visible = False
   Back.SetFocus
   'New1.SetFocus ' added on 9499
    Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Public Sub cleartxt()
'com_mno = " "
'Com_cno.Text = " "
lbl_typ = " "
'Com_dno.Text = " "
Text22.Text = " "
lbl_age = " "
lbl_cat = " "
lbl_sex = " "
LBL_NAME = " "
lbl_dname = " "
'txt_c_cr = " "
lbl_tot_amt = " "
'' SANDHYA PROTOCOL PATIENT
nettot = " "
prottot = " "
PTOT = 0
NTOT = 0
'' END OF PROTOCOL
lbl_dcode = " "
d_u_name = "  "
me_cno.Mask = ""
me_cno.Text = ""
lbl_cr_card_no = " "
lbl_tr_no = " "
me_mno.Mask = "  "

'Set DBGrid1.DataSource = Nothing '24-AUG
'DBGrid1.SetFocus
DBGrid1.row = 1
DBGrid1.col = 0

DBGrid1.Refresh
DBGrid1.ClearFields
DBGrid1.Refresh
DBGrid1.Rebind
DBGrid1.Refresh
DBGrid1.row = 1
For SI = 0 To DBGrid1.Columns.count - 1
DBGrid1.col = SI
DBGrid1.Columns(SI).Text = ""
Next
DBGrid1.Refresh
DBGrid1.row = 1
DBGrid1.col = 0
'DBGrid1.SetFocus


'lbl_name.Caption = ""
'lbl_age.Caption = ""
'lbl_cr_card_no.Caption = ""
'lbl_sex.Caption = ""
'lbl_tr_no.Caption = ""
'lbl_cat.Caption = ""
'lbl_typ.Caption = ""
'lbl_dno = " "

txt_doc_sec_name = " "
txt_rem = " "
txt_unit_doc_name = " "
txt_req_no.Text = ""
INNC = False

cbo_reqno.clear
'txt_unit_doc_name.TabStop = True
End Sub
Sub commit_Click()


Dim iRow As Integer
Dim zconfirm As Integer
Dim zvalidid As Variant
Dim ztotser As Long
Dim i
Dim SR_AMT As Long
FLGTA = 0
FLGCR = 0
ztotser = 0
AmtNotEq = 0

'' BIJESH DUPLICATE MEMOS

memoduplicate = False
' END

'********************************20499
'If Len(Trim(me_mno.Text)) = 0 Then ''19599
'    MsgBox "Memo number can't be empty !", vbCritical
'    me_mno.SetFocus
'    Exit Sub
'End If ''19599

If EmptyGrid = 0 Then
MsgBox "ATTEMPT TO SAVE EMPTY MEMO"
Exit Sub
End If
'********************************20499
''''''added by charulata for smart card patient
If Trim(lbl_typ.Caption) = "CP" Then
    Dim CSNR As String
    Dim PIN As String
    Dim cardpin As String '''to store card pin
    Dim notimes As Integer
    Dim ANS As String
    Dim cardsrno As String
    
    If scard = True Then
    Set PCARD = New PISWriteCard.ClsWriteCard
    If Trim(ValidateCard(me_cno.FormattedText)) <> "" Then
        cardsrno = PCARD.GetCardInfo(Strkey, PINFO)
        If PINFO.PatientCaseNumber = "" Then
            MsgBox ("Place the card on the Reader")
            Exit Sub
        Else
            If Trim(ValidateCard(me_cno.FormattedText)) <> Trim(PCARD.GetCSNR(Strkey, PINFO)) Then
                MsgBox ("Not a Vaild Card.")
                Exit Sub
            End If
        End If
        
        
        If GETSMARTCARD(me_cno.FormattedText, PCARD.GetCSNR(Strkey, PINFO)) = "1" Then
        
            cardsrno = Trim(PCARD.GetCSNR(Strkey, PINFO))
            lbl_csnr.Caption = cardsrno
            If Val(lbl_deposit.Caption) < Val(lbl_tot_amt.Caption) Then
                MsgBox ("Service Amount is greater than Smart-Card balance amount.")
                Exit Sub
            End If
            'MsgBox ("Place the card on the reader.")
            scard = True
            notimes = 0
        
        
PIN:
            DisplayPole ("Enter your Pin                          ")
            PIN = Val(PasswordInputbox("Enter Your Pin", "PIN"))
        
            If Trim(PIN) <> "0" Then
                PINFO.PinNo = Encrypt(Trim(PIN))
                If PCARD.VerifyPin(Strkey, PINFO) = "" Then
                    If notimes = 5 Then
                        MsgBox ("Your account has been locked. As you have made three consecutive erroneous entries.")
                        LOCKACCOUNT (cardsrno)
                                       
                        Exit Sub
                    End If
                    MsgBox ("Pin number is not correct.")
                    notimes = notimes + 1
                    GoTo PIN
                Else
              
                End If
            Else
                MsgBox ("Pin cannot be blank.")
                ANS = MsgBox("Do you want to cancel?", vbYesNo)
                    If ANS = vbNo Then
                        GoTo PIN
                    Else
                        Exit Sub
                    End If
                        
            
            End If
        Else
            scard = False
        
        End If
        Set PCARD = Nothing
    End If
End If
End If
'''''upto here '''''''


'26/10/2005--BIJESH
'PatchUp for checking if the same Service code has been generated for the patient
'in the same day.

'This query is used to get the AS400 Date and Time
Set ZAS400DTTM = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1,CURRENT TIME TM1 FROM PATOOLSLIB.MA107") 'added for date
If rs2flag = False Then
    OPQ(0) = me_cno.FormattedText
    Set rs2 = OPQ.OpenResultset(1, 3)
    rs2flag = True
Else
    OPQ(0) = me_cno.FormattedText
    rs2.MoreResults
    rs2.Requery
End If

'''' SANDHYA DATED 12.3.2013 TO CHECK IF REGISTRATION IS ALREADY DONE
If Trim(UserData(0, iRow)) = "A001" Then
For iRow = 0 To mTotalRows&
  If Len(Trim(UserData(0, iRow))) > 0 Then '''''28041999
    If Not IsEmpty(UserData(0, iRow)) Then
       zgetSerA001(0) = CASENO
       zgetSerA001(1) = Trim(UserData(0, iRow))
       zrsgetSerA001.Requery
            
       If zrsgetSerA001(0) > 0 Then
           MsgBox ("SERVICE CODE FOR REGISTRATION ALREADY GENERATED, CANNOT GENERATE TWICE ")
          
           FLGCR = 0
           btn_vis (0)
           Exit Sub
       End If
          

       
    End If
  End If
Next iRow
End If


''' ADDED BY SHRADDHA FOR DOUBLE MEMO CHECK FOR CR/TA PATIENT
''' DATED 30.8.2010
If rs2("PAT_TYPE") = "CR" Or rs2("PAT_TYPE") = "TA" Then
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
       
        
       '*****************************
       
       If zrsgetSercodes(0) > 0 Then
       
        If rsser_ref.RowCount = 1 Then
            memoduplicate = False
        Else
               ''' ADDED BY SANDHYA DATED 16.8.2010
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
            op_memo_d1.Hide
            Frmpass.ZOrder
            DoEvents
            Frmpass.show MODAL
          End If
       End If
 ''''' commented for checking duplicate ser_code
 
       
'          I = MsgBox("The Service code  " & Trim(UserData(0, iRow)) & " has already been generated " & zrsgetSercodes(0) & " times with total " & zrsgetSercodes(1) & " quantities for this Patient Today..." & Chr(13) & "Do You Wish to Proceed.... ????", vbYesNo + vbDefaultButton2, "Confirmation Required..??")
'          If I = vbNo Then
'             memoduplicate = True
'          End If
       End If
    End If
  End If
Next iRow
End If ' (DOUBLE MEMO CHECK END)
If memoduplicate = True Then
   Exit Sub
Else


' ****** CHECK THIS PARA SANDHYA
'********************************20499
'If rs2flag = False Then
'    OPQ(0) = me_cno.FormattedText
'    Set rs2 = OPQ.OpenResultset(1, 3)
'    rs2flag = True
'Else
'    OPQ(0) = me_cno.FormattedText
'    rs2.MoreResults
'    rs2.Requery
'End If

' **** SANDHYA

If Len(Trim(lbl_tot_amt.Caption)) <> 0 Then
If rs2("PAT_TYPE") = "CR" And BPPATFLAG = False Then
    If rs_crflag = False Then
        Ta114Qr(0) = CASENO
        Set rs_cr = Ta114Qr.OpenResultset(1, 3)
        rs_crflag = True
    Else
        Ta114Qr(0) = CASENO
        rs_cr.MoreResults
        rs_cr.Requery
    End If
    ''''Set rs_cr = rdb.OpenResultset("select * from ta114 where case_no = '" & CASENO & "'", 2, 3)
        If rs_cr("max_amt") <> 0 Then
                    If Len(Trim(rs_cr("SER_AMT"))) = 0 Then
                    SR_AMT = 0
                    Else
                    SR_AMT = rs_cr("ser_amt")
                    End If

                 If (SR_AMT + CDbl(lbl_tot_amt)) > (0.9 * rs_cr("MAX_AMT")) Then
                    MsgBox "SERVICE AMOUNT EXCEEDS 90% OF MAXIMUM AMOUNT ; MEMO CANNOT BE ISSUED "
                    MsgBox "COMPANY BALANCE AMOUNT = " & (rs_cr("MAX_AMT") - SR_AMT)
                    FLGCR = 0
                    btn_vis (0)
                    Exit Sub
                 Else
                    FLGCR = 1

                 End If
        
        Else
                 If rs_cr("date_max") < CDate(me_date) Then
                       MsgBox "VALID DATE EXCEEDED ...."
                       FLGCR = 0
                       btn_vis (0)
                       Exit Sub
                End If
                MsgBox "MEMO IS TO BE SAVED"
        End If
   
ElseIf rs2("PAT_TYPE") = "TA" Then
     ' Padmakar - (TMH-ACTREC) joint company account details
     Call TMH_ACTREC(CASENO)

        If rs_trflag = False Then
            Ta113Qr(0) = CASENO
            Set rs_tr = Ta113Qr.OpenResultset(1, 3)
            rs_trflag = True
        Else
            Ta113Qr(0) = CASENO
            rs_tr.MoreResults
            rs_tr.Requery
        End If
        ''''Set rs_tr = rdb.OpenResultset("select * from ta113 where case_no = '" & CASENO & "'", 2, 3)
            'If Not rs_tr.EOF Then
                    If Len(Trim(rs_tr("SER_AMT"))) = 0 Then
                        SR_AMT = 0
                    Else
                        SR_AMT = CDbl(rs_tr("ser_amt"))
                    End If
            'End If
                    lbl_tot_amt.Refresh
                     iRow = 0
             If UserData(0, iRow) <> "Z100" Then
                If (SR_AMT + CDbl(lbl_tot_amt)) > (0.9 * rs_tr("NET_AMT")) Then
                    MsgBox "MEMO AMOUNT EXCEEDS 90% OF MAXIMUM AMOUNT ; MEMO CANNOT BE ISSUED "
                    ztotser = rs_tr("ser_amt") + CDbl(lbl_tot_amt.Caption)
                    MsgBox "TRUST BALANCE AMOUNT = " & (rs_tr("NET_AMT") - (ztotser))
                    
                   '***************************
                   '20/02/2002-Parag
                   'Even if the amount exceeds 90% of the total deposited
                   'amount but the balance after reducing the new memo amt
                   'is more than Rs100 the memo can be saved.
                   lbl_tot_amt.Refresh
                   If (rs_tr("NET_AMT") - (ztotser)) >= 100 Then
                       FLGTA = 1
                       MsgBox "MEMO IS TO BE SAVED"
                   Else
                    '----------------------------------
                    'Even if 90% amount is exceeded for Group 3 (Accounts Group)
                    'for users listed in matchup function can save the memo
                    'on the basis of a sanction given by Authorised personnel.
                    '06/08/2001-Parag
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
                If (ztotser) > (rs_tr("NET_AMT")) Then
                    MsgBox " Service amount exceeds balance amount", vbInformation, "Amount Exceeds"
                    FLGTA = 0
                    btn_vis (0)
                    Exit Sub
                Else
                    FLGTA = 1
                    MsgBox "MEMO IS TO BE SAVED"
                End If
              
            End If
    End If
End If
'V001 PATIENT DEPOSIT STARTS
      If rs2("PAT_TYPE") = "CP" Then
      
        TA116Qr(0) = CASENO
        Set rs_tr = TA116Qr.OpenResultset(2, 3)
        
        If rs_tr.RowCount = 1 Then
           PATDEPOSIT
             If STAT = "0" Then
               CANCEL_Click
               Exit Sub
             End If
            
        End If
      End If
    'V001 PATIENT DEPOSIT ENDS
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
      prottot.Refresh
      If (SR_AMT + CDbl(PTOT)) > (0.75 * PROTCASE_RS("TOTALAMT")) Then
                    MsgBox "MEMO AMOUNT EXCEEDS 75% OF MAXIMUM AMOUNT ; MEMO CANNOT BE ISSUED "
                    ztotser = PROTCASE_RS("ser_amt") + CDbl(prottot)
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
'     If (ztotser) > (PROTCASE_RS("SANC_AMT")) Then
'         MsgBox " Service amount exceeds balance amount", vbInformation, "Amount Exceeds"
'         FLGTA = 0
'         btn_vis (0)
'         Exit Sub
'     Else
'        FLGTA = 1
'        MsgBox "MEMO IS TO BE SAVED"
'     End If
              
   End If
 End If
'' protocol patient ends
'On Error GoTo MSG:
i = MsgBox("Do you want to save? ", vbOKCancel, "SAVE BOX")
''added by mono
 'If Trim(cbo_reqno.Text) <> "" Then
 
 If cbo_reqno.ListCount <> 0 Then
    MsgBox " Please select all the Procedure from the Combo Box !!", vbInformation, "MEMO"
    cbo_reqno.SetFocus
    Exit Sub
 End If
If i = vbOK Then
   add_edit
End If
Me.Visible = True
 btn_vis (0)
' CLEARTXT '5599
' Refresh1 '5599
 Enblbox (0) '''5599
 '''New1.SetFocus
'print1.Value = True
print1.SetFocus
'msg:
 '       MsgBox Error(Err)
End If '  (BIJESH DUPLICATE MEMOS)
End Sub

Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
'Dim iCol As Integer
'
'mTotalRows& = mTotalRows& + 1
'ReDim Preserve UserData(MAXCOLS - 1, mTotalRows& - 1)
'NewRowBookmark = mTotalRows& - 1
'
'For iCol = 0 To UBound(UserData, 1)
'    If Not IsNull(RowBuf.Value(0, iCol)) Then
'        UserData(iCol, mTotalRows& - 1) = RowBuf.Value(0, iCol)
'    Else
'
'       UserData(iCol, mTotalRows& - 1) = DBGrid1.Columns(iCol).DefaultValue
'    End If
'Next iCol
'
End Sub

Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

For iRow = Bookmark + 1 To mTotalRows& - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows& = mTotalRows& - 1

End Sub

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
For iRow = 0 To RowBuf.RowCount
    'If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
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
Dim i As Integer
Dim str As String
Screen.MousePointer = 0
'STR = "select * from TOOLSADMN.scrn_auth1 where TRIM(USR_GRP) = ? AND TRIM(SCRN_ID) = ?"
'Set Scrn_AuthQr = rdb.CreateQuery("Scrn_AuthQ", STR)
'LVL_BTN_ENABL '**1/7/99
''NEW_LVL_BTN
Delete.Visible = False
db_flg = 0
CANCELFLG = False
flag1 = 0
AmtNotEq = 0
memoduplicate = False
'Main
'INNC = False
'ReDim UserData(0 To 5, 0 To 20)
ReDim UserData(0 To 5, 0 To 12)
'mTotalRows& = 19
mTotalRows& = 12
Dim oldcnt As Integer, newcnt As Integer

'com_mno.Visible = False
lbl_dno.Visible = False
lbl_cno.Visible = False
flag = 0

flex1.Visible = False
flex2.Visible = False


'***** Added by Shraddha on 27/10/12 *****
str = "SELECT * FROM DISLIB.DISICU where IPADD=? and m1 in ('I1','M1','M2','M3','M4','M5','M6','M7','MH','M8')"
Set SEL_DISICU_PM_QR = rdb.CreateQuery("", str)
'*****************************************

a = 1
 'This query gets the doctor code from DOCVIEW
 Set RS33QR = rdb.CreateQuery("IPRS33Q", "SELECT DOC_CODE FROM PATOOLSLIB.DOCVIEW WHERE UD_NAME = ?") 'BILLOPT 61299 ADDED FETCHING FROM VIEW
 
 Set RS33QRCD = rdb.CreateQuery("IPRS33QCD", "SELECT UD_NAME FROM PATOOLSLIB.DOCVIEW WHERE DOC_CODE = ?") 'BILLOPT 61299 ADDED FETCHING FROM VIEW
If opmemoflag = False Then

''29/10/03--MONO
 ''******FOR BP CROSSMATCH**********
 str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=? AND PRO_CODE='H002'"
 Set SELH002_QR = rdb.CreateQuery("SELH002_STR", str)
 
str = "SELECT COLLECTION FROM DISLIB.DMULPROCE WHERE REQNO=? AND PRO_CODE='H002'"
Set SELCOLLECTION_QR = rdb.CreateQuery("SELCOLLECTION_STRRR", str)


''******FOR BP CROSSMATCH**********

str = "select * from PATOOLSLIB.MA106 WHERE DEPT_NAME <> ?"
Set Ma106Qr = rdb.CreateQuery("Ma106Q", str)
'Ma106Qr(0) = " "
'Set rs12 = Ma106Qr.OpenResultset(2, 3)
''''Set rs12 = rdb.OpenResultset("select * from MA106", rdOpenKeyset, 1)
'rs12.MoveFirst
'While Not rs12.EOF
'    'Com_dno.AddItem rs12("dept_code")
'    Com_dno.AddItem rs12("dept_name") '***31399
'    ''Com_dno.ItemData(Com_dno.NewIndex) = rs12("DEPT_CODE")
'    rs12.MoveNext
'Wend

'Set rs7 = rdb.OpenResultset("select * from TA1071", 2, 3)
'''Set rs7 = rdb.OpenResultset("select * from TA1071 where 1 = 2", 2, 3) '''C221199

'Me.Show 'WHY IT IS REQUIRED
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
    
    With DBGrid1
    
       ' .Columns(0).Width = 1000
        .Columns(0).Width = 3700
        .Columns(0).Caption = "CODE/REQ.NO/REQ.DATE/REQ.CATG"
        .Columns(0).Visible = True
        .Columns(0).Button = True
        .Columns(1).Width = 4200
        .Columns(1).Caption = "SERVICES/PROCEDURES/INVESTIGATIONS"
        .Columns(1).Visible = True
        .Columns(2).Width = 510
        .Columns(2).Caption = "QTY"
        .Columns(2).Visible = True
        .Columns(3).Caption = "AMOUNT"
        .Columns(3).Visible = True
        .Columns(4).Width = 1500
        .Columns(4).Caption = "DOCTOR CODE"
        .Columns(4).Visible = True
        .Columns(5).Width = 3000
        .Columns(5).Caption = "DOCTOR NAME"
        .Columns(5).Visible = True
    End With
'''Set RS14 = rdb.OpenResultset("select ud_name from MA1031", rdOpenKeyset)
'''Set RS16 = rdb.OpenResultset("select doc_name from MA102", rdOpenKeyset)
str = "select ud_name,U_D from PATOOLSLIB.MA1031 WHERE UD_NAME <> ?"
Set Ma1031Qr = rdb.CreateQuery("Ma1031Q", str)
'Ma1031Qr(0) = " "
'Set RS17 = Ma1031Qr.OpenResultset(2, 3)

''''Set RS17 = rdb.OpenResultset("select ud_name,U_D from MA1031", rdOpenKeyset)
'Set Ma102Qr = rdb.CreateQuery("Ma102Q", "select * from ma102 where doc_code <> ? ORDER BY DOC_CODE")
Set Ma102Qr = rdb.CreateQuery("Ma102Q", "select * from PATOOLSLIB.DOCVIEW where doc_code <> ? ORDER BY DOC_CODE")
'Ma102Qr(0) = 0
'Set rs_dname = Ma102Qr.OpenResultset(2, 3)
''''Set rs_dname = rdb.OpenResultset("select * from ma102 ORDER BY DOC_CODE", rdOpenKeyset)
Set MA102DOC = rdb.CreateQuery("Ma102DOCQ", "select DOC_NAME from PATOOLSLIB.MA102 where doc_code = ? ")
MA102DOC(0) = 0
Set OPQ = rdb.CreateQuery("OPQ1", "SELECT first_name, mid_name, last_name, sex, pat_catg, pat_reg_ty, pat_type, ud_name, birth_date, consult_dt, u_d  FROM PATOOLSLIB.TA101 WHERE CASE_NO = ?") '''A251199
'OPQ(0) = "BN/00001"
'Set rs2 = OPQ.OpenResultset(2, 3)

'STR = "INSERT INTO RISLIB.TA107 (memo_no, io_memo, memo_date, dept_code, case_no, MEMO_CATG, d_s_name, rem,  doc_code, amt_rcvd, memo_amt) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
'Set Ta107_InsrtQr = rdb.CreateQuery("Ta107_InsrtQ", STR)
'
'STR = "INSERT INTO RISLIB.TA1071 (MEMO_NO, SER_CODE, QTY, SER_AMT, DOC_CODE) VALUES (?, ?, ?, ?, ?)"
'Set Ta1071_InsrtQr = rdb.CreateQuery("Ta1071_InsrtQ", STR)

'str = "INSERT INTO RISLIB.TA107 (memo_no, io_memo, memo_date, dept_code, case_no, pat_CATG, d_s_name, rem,  doc_code, amt_rcvd, memo_amt,REQ_NO,REQ_DATE) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?)" ' comm on 24/7
'******

''STR = "INSERT INTO PATOOLSLIB.TA107 (memo_no, io_memo, memo_date, dept_code, case_no, MEMO_CATG, d_s_name, rem,  doc_code, amt_rcvd, memo_amt) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
''Set Ta107_InsrtQr = rdb.CreateQuery("Ta107_InsrtQ", STR)
''
''STR = "INSERT INTO PATOOLSLIB.TA1071 (MEMO_NO, SER_CODE, QTY, SER_AMT, DOC_CODE,REQ_NO,req_date,req_catg) VALUES ( ?,?,?, ?,?,?, ?, ?)"
''Set Ta1071_InsrtQr = rdb.CreateQuery("Ta1071_InsrtQ", STR)

str = "INSERT INTO PATOOLSLIB.TA107 (memo_no, io_memo, memo_date, dept_code, case_no, MEMO_CATG, d_s_name, rem,  doc_code, amt_rcvd, memo_amt,SHIFLAG) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?)"
Set Ta107_InsrtQr = rdb.CreateQuery("Ta107_InsrtQ", str)

''******on 26/06/03
str = "INSERT INTO PATOOLSLIB.TA1071 (MEMO_NO, SER_CODE, QTY, SER_AMT, DOC_CODE,REQ_NO,req_date,req_catg,PROF_AMT,HOSP_AMT,DEPT_CODE,fld1) VALUES ( ?,?,?, ?,?,?, ?, ?,?,?,?,?)"
Set Ta1071_InsrtQr = rdb.CreateQuery("Ta1071_InsrtQ", str)
'******

'***** Shraddha 27/07/10 *****
str = "SELECT * FROM DISLIB.DISWARD WHERE IPADD=?"
Set SELDISWARD_QR = rdb.CreateQuery("SELDISWARD_STR", str)
'*****************************


str = "update PATOOLSLIB.ta113 set ser_amt = ser_amt + ? where (case_no = ?)"
Set Ta113_UpdtQr = rdb.CreateQuery("Ta113_UpdtQ", str)
 
str = "update PATOOLSLIB.ta114 set ser_amt = ser_amt + ? where case_no = ?"
Set Ta114_UpdtQr = rdb.CreateQuery("Ta114_UpdtQ", str)
 
str = "SELECT doc_code FROM PATOOLSLIB.DOCVIEW WHERE TRIM(ud_name) = ?"
Set Doc_ViewQr = rdb.CreateQuery("Doc_ViewQ", str)
'Doc_ViewQr(0) = "J J VYAS"
'Set rs33 = Doc_ViewQr.OpenResultset(2, 3)

str = "select * from PATOOLSLIB.ta113 where case_no = ?"
Set Ta113Qr = rdb.CreateQuery("Ta113Qr", str)

str = "select * from PATOOLSLIB.ta114 where case_no = ?"
Set Ta114Qr = rdb.CreateQuery("Ta114Qr", str)

'STR = "select * from PATOOLSLIB.TA107 where memo_no = ? AND io_memo = ? "
str = "select C_CR,MEMO_NO,DEPT_CODE,CASE_NO,MEMO_DATE,DOC_CODE,MEMO_CATG,MEMO_AMT,REM,D_S_NAME,DIS_FLAG,IO_MEMO,AMT_RCVD from PATOOLSLIB.TA107 where memo_no = ? AND io_memo = ? "
Set Ta107Qr = rdb.CreateQuery("Ta107Q", str)
 
'STR = "select * from RISLIB.TA107 where memo_no = ?"
'Set Ta107Qr = rdb.CreateQuery("Ta107Q", STR)

str = "select * from PATOOLSLIB.TA1071 where memo_no = ?"
Set Ta1071Qr = rdb.CreateQuery("Ta1071Q", str)

str = "select case_no from PATOOLSLIB.ta105 where case_no = ? AND DISC_DATE IS NULL AND ADMS_DATE IS NOT NULL"
Set InPt_ChkQr = rdb.CreateQuery("InPt_ChkQ1", str)

str = "select memo_no,D_S_NAME,IO_MEMO from PATOOLSLIB.ta107 where memo_no = ?"
Set Exist_ChkQr = rdb.CreateQuery("Exist_ChkQ", str)

str = "select * from PATOOLSLIB.MA101 where ser_code = ?"
Set Ma101Qr = rdb.CreateQuery("Ma101Q", str)
'Ma101Qr(0) = "G101"
'Set RS18 = Ma101Qr.OpenResultset(2, 3)

str = "select ser_code, ser_det from PATOOLSLIB.MA101 where ser_det <> ? order by ser_code"
Set AllMa101_Qr = rdb.CreateQuery("AllMa101_Q", str)
str = "DELETE FROM PATOOLSLIB.TA107 WHERE MEMO_NO = ?"
Set DELTA107 = rdb.CreateQuery("DELTA107Q", str)
str = "DELETE FROM PATOOLSLIB.TA1071 WHERE MEMO_NO = ?"
Set DELTA1071 = rdb.CreateQuery("DELTA1071Q", str)

'STR = "SELECT * FROM RISLIB/TA1071 WHERE MEMO_NO=? "
'SELTA1071_QR=RDB.CreateQuery (
 btn_vis (0)
 Enblbox (0) '''5599
 
 'ADDED FOR RIS 11/05/2001
   'Set rq_reqdet = rdb.CreateQuery("rq_reqdet_id", "select reqnno,reqndt,patnstat,dept_code,pat_catg,ser_code,refbydoc from rrequisitn where caseno = ? and patnstat = ? and reqnno = ? and dept_code = ? and ((reqnstat is null)  or reqnstat = 'A')")
   Set rq_reqdet = rdb.CreateQuery("rq_reqdet_id1", "select reqnno,reqndt,patnstat,dept_code,pat_catg,ser_code,refbydoc,RECORDNO from Rislib.rrequisitn where caseno = ? and patnstat = ? and reqnno = ?  and ((reqnstat is null)  or (reqnstat='P'  and (RECORDNO=1 and RECPTNO='VPP')))")
   Set rq_updtreqdet = rdb.CreateQuery("rq_updtreqdet_id1", "update Rislib.rrequisitn set memodt = ?,memono = ?,reqnstat = ? where caseno = ? and reqnno = ?  and ((reqnstat is null) or reqnstat = 'A')")
  Set patndt_qr = rdb.CreateQuery("patndt_qr", "select patnstat from rislib.rrequisitn where caseno=? and reqnno=?")
   ' Set rq_updtreqdet = rdb.CreateQuery("rq_updtreqdet_id", "update rislib.rrequisitn set memodt = ?,memono = ?,reqnstat = ? where caseno = ? and reqnno = ? and dept_code = ? and ((reqnstat is null) or reqnstat = 'A')")
  ' Set RQ_REQDETREQNO = rdb.CreateQuery("RQ_REQDETREQNO_ID", "SELECT reqnno FROM rislib.rrequisitn WHERE CASENO = ? AND patnstat = ? and dept_code = ? AND ((reqnstat is null)  or reqnstat = 'A')")
   ' if u want all departments requisitions remove the comments and add comment for the above chandu
  Set RQ_REQDETREQNO = rdb.CreateQuery("RQ_REQDETREQNO_ID", "SELECT reqnno FROM Rislib.rrequisitn WHERE CASENO = ? AND patnstat = ?  AND ((reqnstat is null)  or reqnstat = 'A')")
 'ADDED FOR RIS 11/05/2001
 
  sql = "select pat_catg from Rislib.rrequisitn where caseno = ?"
  Set requisitn_qr = rdb.CreateQuery("requ_qr11", sql)
'  requisitn_qr(0) = "BN/00000"
'  Set requisitn_rs = requisitn_qr.OpenResultset(2, 3)
  
  Set REQUISITN_UPDT = rdb.CreateQuery("REQUISITN_UPDT11", " UPDATE Rislib.RREQUISITN SET RECPTNO=?,RECPTDT=?,  REQNSTAT=? WHERE REQNNO=?")
  Set disreq_updt = rdb.CreateQuery("disicuup", "update rislib.rrequisitn set reqnstat = ? where reqnno =?")
  
  sql = "select req_no from PATOOLSLIB.ta1071 where memo_no=?"
  Set RQ_reqdet1 = rdb.CreateQuery("rq_reqdet111", str)
  RQ_reqdet1(0) = "M-000000"
  Set rs_reqdet1 = RQ_reqdet1.OpenResultset(2, 3)
Set UPDTA107_QR = rdb.CreateQuery("1UPDTA107QR1", "UPDATE PATOOLSLIB.TA107 SET AMT_RCVD=? WHERE MEMO_NO=?")
 
 sql = "SELECT YEAR1,SERIES FROM PATOOLSLIB.AUTO_MEMO WHERE YEAR1 =? "
 Set RQ_SERIES = rdb.CreateQuery("1RQ_SERIES1", sql)
 
 'Set RS_SERIES = RQ_SERIES.OpenResultset(2, 3)
 
 'modified by v.s.k. on 04012012
 'Set rq_updtSERIES = rdb.CreateQuery("1rq_updtSERIESmemosrno_id", "update PATOOLSLIB.AUTO_MEMO set SERIES = ? WHERE YEAR1 = ? ")  ' ADDED 11/11/2001
 Set rq_updtSERIES = rdb.CreateQuery("1rq_updtSERIESmemosrno_id", "update PATOOLSLIB.AUTO_MEMO set SERIES = ? WHERE YEAR1 = ? and SERIES<?")  ' ADDED 11/11/2001
 
 sql = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? "
 Set SELRREQQR = rdb.CreateQuery("SELRREQQROP", sql)
 '--------------------------------------------------------------------------------
 '06/04/2002-Parag
 'if a receipt is generated against the supplied memo number
 'the details of receipt are displayed in "Remarks" column.
 zstr = "SELECT * FROM PATOOLSLIB.TA110 WHERE PRN_FLAG='P' AND REC_NO IN (SELECT REC_NO FROM PATOOLSLIB.TA1101 WHERE M_B_NO=?)"
 Set zmemorcpt = rdb.CreateQuery("zmemorcpt1", zstr)
    opmemoflag = True
    
 sql = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
 Set SELDMULPROCEQR = rdb.CreateQuery("SELDMULPROCEQROP", sql)
   
 sql = "SELECT * FROM DISLIB.DMOLEPROC WHERE REQNO=?"
 Set SELDMOLEPROCQR = rdb.CreateQuery("SELDMOLEPROCEQ", sql)
 
 zstr = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
 Set zgetservice = rdb.CreateQuery("zgetservOPM1", zstr)
 
 '**** MAPING FROM SERVICE CODE AND DOCTOR CODE
 
     'Get the details from mapping table : SHI1.SERVICED
    zstr = "SELECT * FROM PATOOLSLIB.SERVICED WHERE SER_MAIN=? ORDER BY SER_CODE"
    Set zmaptable = rdb.CreateQuery("zmaptableqOP", zstr)
    'zmaptable(0) = "A"
    'Set zrsmaptable = zmaptable.OpenResultset(2, 1)
    
    'Update Dept code and Doc Code as per the mapping table (Service Code - Doc Code and Dept code)
    zstr = "UPDATE PATOOLSLIB.TA1071 SET DEPT_CODE =?,DOC_CODE=? WHERE SER_CODE=? AND MEMO_NO=?"
    Set zupdtmap = rdb.CreateQuery("zupdtmapqOP", zstr)
'    zupdtmap(0) = "AAA"
'    zupdtmap(1) = "99999"
'    zupdtmap(2) = "A001"
'    zupdtmap(3) = "M2999999"
    
    '--------------------------------------------------------------------------------
    '18/10/2002-Parag/Sandhya
    'This query is used to get all records from SHI1.patoolslib_ta1071
    zstr = "SELECT * FROM PATOOLSLIB.TA1071 WHERE MEMO_NO=? ORDER BY SER_CODE"
    Set zta1071 = rdb.CreateQuery("zta1071MAPOP", zstr)
'    zta1071(0) = "M999999"
'    Set zrsta1071 = zta1071.OpenResultset(2, 1)
    '--------------------------------------------------------------------------------
    
 '**** MAPING FROM SERVICE CODE AND DOCTOR CODE
    zstr = "select * from rislib.rrequisitn where reqnNo=?"
    Set selreq_qr = rdb.CreateQuery("selreq_qrop", zstr)
 
 End If
 '--------------------------------------------------------------------------------
 ' **** SEKHAR IS ADDED  THESE CODE

 btn_vis (1)
 Dim REQNO As String
' i = pathology.Data1.Recordset.RecordCount
'
' If i <> 0 Then
'
'    pathology.Data1.Recordset.MoveFirst
'
'    While Not pathology.Data1.Recordset.EOF
'        cbo_reqno.AddItem pathology.Data1.Recordset.Fields("reqno")
''        reqno = req_chglen(pathology.Data1.Recordset.Fields("reqno"))
''        CASENO = pathology.Msk_caseno.FormattedText
''        If Trim(lbl_cat.Caption) = "NC" And INNC = True Then
''            rq_reqdet(0) = CASENO
''            rq_reqdet(1) = "IP"
''            rq_reqdet(2) = reqno
''        Else
''            rq_reqdet(0) = CASENO
''            rq_reqdet(1) = "OP"
''            rq_reqdet(2) = reqno
''        End If
''            Set rs_reqdet = rq_reqdet.OpenResultset(1, 3)
''        If rs_reqdet.RowCount = 0 Then
''             MsgBox "No corresponding req. no. record exists or either its reqnstat is M in rrequisitn table"
''        Else
''               ReDim Preserve UserData(0 To 5, 0 To 12)
''              DBGrid1.Col = 0
''              'DBGrid1.row = 0
''              DBGrid1.Text = rs_reqdet("ser_code") & " " & reqno & " " & rs_reqdet("reqndt") & " " & Trim(rs_reqdet("pat_catg"))
''              me_rdate.Text = rs_reqdet("reqndt")
''
''                 lbl_cat = Trim(rs_reqdet("pat_catg"))
''
''              DBGrid1.Col = 4
''              DBGrid1.Text = lbl_dcode
''              DBGrid1.Col = 5
''              DBGrid1.Text = rs_reqdet("REFBYDOC")
''
''          End If
''          DBGrid1.Refresh
''              'DBGrid1_AfterColEdit()
                
    
      ' pathology.Data1.Recordset.MoveNext
'    Wend
' End If
'' sandhya 1.8.2008 protocol patient

    str = "SELECT * FROM PATOOLSLIB.PROTOCASE WHERE CASE_NO = ? AND IRBNO = ? "
    Set PROTCASE_QR = rdb.CreateQuery("protCASE_QR31", str)
    
        
    str = "SELECT * FROM PATOOLSLIB.PROTSERVIC WHERE irbNO = ? and ser_code = ? and (CURRENT DATE BETWEEN FROMDT AND TODT) AND STATUS = 'C' "
    Set PROTSERV_QR = rdb.CreateQuery("PROTSERV_AQR11", str)
    
    
    str = "INSERT INTO PATOOLSLIB.PROTMEMO (MEMO_NO,MEMO_DATE,irbNO,CASE_NO,SER_code,SER_amt,prof_amt,hosp_amt,qty,MEMO_AMT,DEPT_CODE,DOC_CODE,PAT_TYPE,PAT_CATG,REQ_NO,req_date,IO_MEMO,AMT_RCVD,USERID,IPADDR) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set PROTMEMO_Qr = rdb.CreateQuery("PROTMEMO_tQ1", str)
    
    str = "update PATOOLSLIB.PROTOCASE set ser_amt = ser_amt + ? where (case_no = ? AND IRBNO = ? )"
    Set PROTCASE_UPD = rdb.CreateQuery("PROTCASE_Upd1", str)
    
    str = "select * from PATOOLSLIB.PROTMEMO where memo_no = ?"
    Set QR_PRTMEMO = rdb.CreateQuery("QRPRTMEMO13", str)

    'Update Dept code and Doc Code as per the mapping table (Service Code - Doc Code and Dept code)
    zstr = "UPDATE PATOOLSLIB.protmemo SET DEPT_CODE =?,DOC_CODE=? WHERE SER_CODE=? AND MEMO_NO=?"
    Set zupdprotmap = rdb.CreateQuery("zupdPRtmaP1", zstr)
'' protocol patient ends

'Get the details from mapping table : SHI1.SERVICED FOR SER_MAIN+SER_CODE
zstr = "SELECT * FROM patoolslib.SERVICED WHERE SER_MAIN=? AND SER_CODE=? ORDER BY SER_MAIN,SER_CODE,SER_IND"
Set zsercodemap = rdb.CreateQuery("zsercodemapdis", zstr)
zsercodemap(0) = "F"
zsercodemap(1) = "F1"
Set zrssercodemap = zsercodemap.OpenResultset(2, 1)


'Get the details from mapping table : SHI1.SERVICED FOR SER_MAIN+SER_IND
zstr = "SELECT * FROM patoolslib.SERVICED WHERE SER_MAIN=? AND SER_IND=? ORDER BY SER_MAIN,SER_CODE,SER_IND"
Set zserindmap = rdb.CreateQuery("zserindmapdis", zstr)
zserindmap(0) = "F"
zserindmap(1) = "F301"
Set zrsserindmap = zserindmap.OpenResultset(2, 1)

'V001 PATIENT DEPOSIT STARTS
    'This query is used to get the Balance Amount against case_no
    str = "select * from PATOOLSLIB.ta116 where case_no = ?"
    Set TA116Qr = rdb.CreateQuery("Ta116QrSCA", str)
    
    'This query is used to update the service amount by service amount of ta116 with Total Memo Amount
    str = "update PATOOLSLIB.ta116 set ser_amt = ser_amt + ? where (case_no = ?)"
    Set Ta116_UpdtQr = rdb.CreateQuery("Ta116_UpdtQrsca", str)

    

'V001 PATIENT DEPOSIT STARTS
 
'''' ADDED BY SANDHYA DATED 12.3.2013 TO CHECK IF REGISTRATION ALREADY DONE ie A001 SER_CODE ALREADY GENERATED

str = "SELECT COUNT(B.SER_CODE),SUM(B.QTY) FROM PATOOLSLIB.TA107 A,PATOOLSLIB.TA1071 B WHERE A.MEMO_NO=B.MEMO_NO AND A.CASE_NO=? AND B.SER_CODE=?"
Set zgetSerA001 = rdb.CreateQuery("zgetSera0011", str)
zgetSerA001(0) = "TT/00300"
zgetSerA001(1) = "A001"
Set zrsgetSerA001 = zgetSerA001.OpenResultset(2, 1)
 
'26/10/2005--BIJESH
'Querry to get the details of Service Codes against Memos generated for the Patient in the same day.
str = "SELECT COUNT(B.SER_CODE),SUM(B.QTY) FROM PATOOLSLIB.TA107 A,PATOOLSLIB.TA1071 B WHERE A.MEMO_NO=B.MEMO_NO AND A.CASE_NO=? AND A.MEMO_DATE=? AND B.SER_CODE=?"
Set zgetSerCodes = rdb.CreateQuery("zgetSerCodesS", str)
zgetSerCodes(0) = "TT/00300"
zgetSerCodes(1) = "2000-01-01"
zgetSerCodes(2) = "A001"
Set zrsgetSercodes = zgetSerCodes.OpenResultset(2, 1)
'' END
 
'''' added by shraddha
'''' dated 16.8.2010 to check if service code exist in table ser_ref for duplication

str = "select * from patoolslib.ser_ref where SER_CODE=?"
Set qrser_ref = rdb.CreateQuery("", str)
'*************************

''''' ADDED BY SANDHYA DATED 11.5.2011
''''' LISTING DOCTORS WHOSE CESSYN = 4

Set SELMA102QR = rdb.CreateQuery("SELMA102QR11", "SELECT * FROM PATOOLSLIB.MA102 WHERE DOC_CODE <> ? AND trim(DOC_CODE) IN (SELECT EMPCODE FROM PAYROLLLIB.EMPDETAIL WHERE CESSYN = '4') ORDER BY DOC_CODE")
SELMA102QR(0) = 0
Set SELMA102RS = SELMA102QR.OpenResultset(2, 3)

Set SELDOCCODE = rdb.CreateQuery("RSSELDCOD", "SELECT * FROM PATOOLSLIB.MA102 WHERE DOC_NAME LIKE ?")
SELDOCCODE(0) = "%asdjkf%"
Set RSSELDCODE = SELDOCCODE.OpenResultset(2, 3)

'''' ADDED BY SANDHYA DATED 7/7/2012 FOR BP RGJAY PATIENT

str = "select * from PATOOLSLIB.BPPATIENT where case_no = ?"
    Set QRBPPATN = rdb.CreateQuery("", str)

 str = "update PATOOLSLIB.BPPATIENT set ser_amt = ser_amt + ?, NET_AMT=NET_AMT - ?  where (caseno = ?)"
    Set UPDbppatn = rdb.CreateQuery("", str)
    
 str = "INSERT INTO PATOOLSLIB.BPPATMEMO (CASENO,PAT_CATG,MEMONO,MEMO_DT,MEMO_AMT,SCHEM_CODE,SCHEM_AMT,USERID,IPADD,FLAG)VALUES(?,?,?,?,?,?,?,?,?,?)"
    Set INS_BPPATMEMO = rdb.CreateQuery("", str)

'''' ADDED BY PADMAKAR (TMH-ACTREC)
'''' TO UPDATE TRUST BALANCE OF ACTREC-TMH

' Trust Amount
str = "UPDATE PATOOLSLIB.TA113 SET NET_AMT=?, SER_AMT = ? WHERE  CASE_NO = ?"
Set s_UpdTa113 = rdb.CreateQuery("s_UpdTa113", str)

str = "SELECT MAX(MEMO_DATE) FROM PATOOLSLIB.TA107 WHERE CASE_NO=?"
Set selActMaxMemo = rdb.CreateQuery("selAMaxMemo", str)

' Company amount
str = "SELECT MAX(MEMO_DATE) FROM PATOOLSLIB.TA107 WHERE CASE_NO=?"
Set selActMaxMemo = rdb.CreateQuery("selActMaxMemo", str)

str = "UPDATE PATOOLSLIB.TA114 SET SER_AMT = ? WHERE  CASE_NO = ?"
Set updTA115a = rdb.CreateQuery("updTA115a", str)


'******************
cbo_reqno_GotFocus
If cbo_reqno.Text = "" Then
'cbo_reqno.AddItem pathology.SHW_REQNO
End If

'txt_unit_doc_name.Text = "S D BANAVALI"

 '**** SEKHAR IS ADDED THESE CODE
 'MsgBox pathology.txt_bcreqno
 '''' added by sandhya for actrec patient treated at tmh
'''' dated 11.6.2012
actmemflg = False
End Sub

Private Sub DBgrid1_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Errhandler
If ColIndex = 0 Then
    flex1.ZOrder
    AllMa101_Qr(0) = ""
    Set RS = AllMa101_Qr.OpenResultset(1, 1)
    ''''Set rs = rdb.OpenResultset("select ser_code,ser_det from MA101 order by ser_code", rdOpenKeyset, 3)
    While Not RS.EOF
     flex1.AddItem RS("ser_code") + vbTab + RS("ser_det")
     RS.MoveNext
    Wend
ElseIf ColIndex = 4 Then
    flex2.ZOrder
    '''Set rs = rdb.OpenResultset("select doc_code from MA102", rdOpenKeyset, 3)
    rs_dname.MoveFirst
    While Not rs_dname.EOF
        flex2.AddItem rs_dname("doc_code")
        rs_dname.MoveNext
    Wend
End If

If a = 1 Then
  If ColIndex = 0 Then
        flex1.Visible = True
        flex1.SetFocus
        a = 0
  ElseIf ColIndex = 4 Then
        flex2.Visible = True
        flex2.SetFocus
        a = 0
  End If
ElseIf a = 0 Then
      If ColIndex = 0 Then
        flex1.Visible = False
        a = 1
      ElseIf ColIndex = 4 Then
            flex2.Visible = False
            a = 1
      End If
End If

If ColIndex = 0 Then
    flex1.Top = DBGrid1.Top + DBGrid1.RowHeight * (DBGrid1.row + 2)
ElseIf ColIndex = 4 Then
        flex2.Top = DBGrid1.Top + DBGrid1.RowHeight * (DBGrid1.row + 2)
End If
Text22.Text = " PRESS CONTROL KEY TO COME OUT OF DBGRID .... and TO REFRESH TOTAL AMOUNT" '****9499
Exit Sub
Errhandler:
MsgBox Err.Description, vbCritical, "MEMO"
Exit Sub
End Sub

Private Sub DBgrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
DBGrid1.Columns(4).Button = True
If LastCol = 0 Then
    flex1.Visible = False
ElseIf LastCol = 4 Then
    flex2.Visible = False
End If
a = 1
Exit Sub
End Sub
Private Sub flex1_Click()
On Error GoTo MSG
Dim opt As String
Dim opttyp As String '***22/10/98
Dim ncchg As String '***22/10/98

''''Set rs5 = rdb.OpenResultset("select * from MA101 where ser_code = '" & flex1.Text & "' or ser_det = '" & flex1.Text & "'", rdOpenKeyset, 3)
If rs18flag = False Then
    Ma101Qr(0) = flex1.Text
    Set RS18 = Ma101Qr.OpenResultset(1, 3)
    rs18flag = True
Else
    Ma101Qr(0) = flex1.Text
    RS18.MoreResults
    RS18.Requery
End If
''' Sandhya protocol patient starts

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

DBGrid1.Columns(0) = RS18("ser_code")
DBGrid1.Columns(1) = RS18("ser_det")
DBGrid1.Columns(2).Text = 1
DBGrid1.Columns(4) = lbl_dcode
opt = lbl_cat
opttyp = lbl_typ '***22/10/98
ncchg = Left(DBGrid1.Columns(0), 1) '***22/10/98
EmptyGrid = 1 ''ADDED ON 17999
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
    
Case "F "
    If opttyp <> "AF" Then '***22/10/98
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH"))
    Else
    DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("FN_HOSP_CH") - RS18("FN_PROF_CH"))  '***22/10/98
    End If
Case "NC"
DBGrid1.Columns(3) = DBGrid1.Columns(2) * (RS18("NC_CH"))  '''''RRRR23899
'If ncchg <> "I" Then '**22/10/98
'    DBGrid1.Columns(3) = 0
'Else                 '**22/10/98
'    DBGrid1.Columns(3) = DBGrid1.Columns(2) * rs5("C_CH") '**22/10/98
'End If '**22/10/98

'    DBGrid1.Columns(3) = 0
Case Else
    MsgBox "NO MATCH ..."
End Select
'AMT_CAL ("rs5")
flex1.Visible = False
DBGrid1.col = 2
DBGrid1.SetFocus
'DBGrid1.Row = DBGrid1.Row + 1
a = 1
Exit Sub
MSG:  MsgBox Error(Err)

End Sub

Private Sub print1_Click()
'On Error Resume Next
'On Error GoTo ERROR_CODE:

'''ADDED BY CHARULATA FOR PRINT ON THERMAL PRINTER
'If scard = True Then
'    themprint
'    Exit Sub
'End If
'''UPTO HERE

'' SANDHYA PROTOCOL PATIENT STARTS
If PROTFLG = True And NPROTFLG = False Then
  PRTMEMOCR
ElseIf PTOT > 0 And NPROTFLG = True Then
  NPRTMEMOCR
ElseIf NPROTFLG = True Then
opmemocr
End If
'' PROTOCOL PATIENT ENDS


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

If PrintResult% = 0 Then
Back.SetFocus
Exit Sub
Else
GoTo ERROR_CODE1
End If

If flag = 1 Then
    If rs2flag = False Then
        OPQ(0) = CASENO
        Set rs2 = OPQ.OpenResultset(1, 3)
        rs2flag = True
    Else
        OPQ(0) = CASENO 'can be removed 'cause rs2 is already having required record
        rs2.MoreResults
        rs2.Requery
    End If
    ''''Set rs2 = rdb.OpenResultset("select * from TA101 where case_no = '" & CASENO & "'", rdOpenKeyset)
    CR1.SelectionFormula = "{TA107.memo_no} = '" & "M-" & me_mno.Text & "'" ' OR {TA107.memo_no} = '" & com_mno.Text & "' "
ElseIf flag = 2 Then
        If rs2flag = False Then
            OPQ(0) = lbl_cno
            Set rs2 = OPQ.OpenResultset(1, 3)
            rs2flag = True
        Else
            OPQ(0) = lbl_cno 'can be removed 'cause rs2 is already having required record
            rs2.MoreResults
            rs2.Requery
        End If
        ''''Set rs2 = rdb.OpenResultset("select * from TA101 where case_no = '" & lbl_cno & "'", rdOpenKeyset)
        CR1.SelectionFormula = "{TA107.memo_no} = '" & "M-" & M_number & "'"
End If

  If rs2("PAT_TYPE") = "TA" Then
     CR1.Formulas(0) = "cr_tr = '" & rs_tr("TR_ACC_NO") & "'"
     ipcr.Formulas(150) = "PATTP = '" & "TRUST PATIENT" & "'"
   ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = False Then
    CR1.Formulas(0) = "cr_tr = '" & rs_cr("CR_CARD") & "'"
    ipcr.Formulas(150) = "PATTP = '" & "COMPANY PATIENT" & "'"
    
   '''' ADDED BY SANDHYA DATED 7/7/2012 FOR RGJAY PATIENT
    ElseIf rs2("PAT_TYPE") = "CR" And BPPATFLAG = True Then
     CR1.Formulas(15) = "cr_tr = '" & Trim(RSBPPATN("BPCARD_NO")) & "'"
     CR1.Formulas(150) = "PATTP = '" & "RGJAY PATIENT" & "'"
   ElseIf rs2("PAT_TYPE") = "CP" And PATDEPO = True Then
     ipcr.Formulas(15) = "cr_tr = '" & rs_tr("PD_ACC_NO") & "'"
     ipcr.Formulas(150) = "PATTP = '" & "DEPOSIT PATIENT" & "'"
   Else
    CR1.Formulas(0) = "cr_tr = '" & "    " & "'"
   End If
   

If Label12.Caption = "UNIT NAME : " Then
     CR1.ReportFileName = "C:\PB121298\Op1.rpt"
   '  CR1.ReportFileName = "C:\PB121298\PARU.rpt"
     
ElseIf Label12.Caption = "DOCTOR NAME : " Then
If flag = 2 Then
    If rs10flag = False Then
        Ta107Qr(0) = "M-" & M_number.Text
        Set rs10 = Ta107Qr.OpenResultset(2, 3)
    Else
         Ta107Qr(0) = "M-" & M_number.Text
        Set rs10 = Ta107Qr.OpenResultset(2, 3)
    End If
Else
    If rs10flag = False Then
        Ta107Qr(0) = "M-" & me_mno '"M-" & M_number.Text
        Set rs10 = Ta107Qr.OpenResultset(2, 3)
    Else
        Ta107Qr(0) = "M-" & me_mno '"M-" & M_number.Text
        Set rs10 = Ta107Qr.OpenResultset(2, 3)
    End If
End If
    
    If rs10.RowCount <> 0 Then
        MA102DOC(0) = Trim(rs10("DOC_CODE"))
        Set rs10 = MA102DOC.OpenResultset(1, 1)
        If rs10.RowCount <> 0 Then
    
         CR1.Formulas(1) = "DOC_NAME = '" & rs10("DOC_NAME") & "'"
        End If
         CR1.ReportFileName = "C:\PB121298\OP.rpt"
      Else
      MsgBox "DOCTOR CODE WAS NOT ASSIGNED"
      Exit Sub
    End If
End If
'CR1.Action = 1
j = CR1.Destination
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
End Sub

Private Sub txt_unit_doc_name_GotFocus()
'If (CASENO = "ZA/99999") Or (CASENO = "ZB/99999") Or (CASENO = "ZD/99999") Or (CASENO = "ZF/99999") Then
' If lbl_cat = "C" Or lbl_cat = "NC" Then
'  'NO OPERATION TO BE DONE
'  ' OTHERWISE GET DOC NAME
'
' Else
'       op_memo_d1.Hide
'       BLOODCOMBO.ZOrder
'        DoEvents
'       BLOODCOMBO.Left = 500
'       BLOODCOMBO.Top = 2000
'       BLOODCOMBO.Show vbModal
'
'       DoEvents
'      ' BLOODCOMBO.BLOODDOC.SetFocus
'       txt_unit_doc_name.Text = BLOODCOMBO.BLOODDOC.Text
'       op_memo_d1.Show
' End If
'
'End If
End Sub

Private Sub txt_unit_doc_name_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'Or KeyAscii = 9 Then
KeyAscii = 0
    'txt_unit_doc_name_LostFocus
 End If
End Sub

Private Sub txt_unit_doc_name_LostFocus()
On Error GoTo MSG:
   If NoCaseFlag = 1 Then '****3499
        'me_cno.SetFocus '****3499
        Exit Sub     '****3499
   End If      '******3499
'*************************
'If (CASENO = "ZA/99999") Or (CASENO = "ZB/99999") Or (CASENO = "ZD/99999") Or (CASENO = "ZF/99999") Then
' If lbl_cat = "C" Or lbl_cat = "NC" Then
'  'NO OPERATION TO BE DONE
'  ' OTHERWISE GET DOC NAME
'
' Else
'       'op_memo_d1.Hide
'       BLOODCOMBO.ZOrder
'        DoEvents
'       BLOODCOMBO.Left = 500
'       BLOODCOMBO.Top = 2000
'       BLOODCOMBO.Show vbModal
'
'       DoEvents
'      ' BLOODCOMBO.BLOODDOC.SetFocus
'      ' BLOODCOMBO.BLOODDOC.SetFocus
'      ' txt_unit_doc_name.Text = BLOODCOMBO.BLOODDOC.Text
'       op_memo_d1.Show
'
'
' End If
'
'End If
'**** to get the doc code
''d_u_name = txt_unit_doc_name
''If Trim(d_u_name) <> "" Then
''MsgBox "HERE IN at txt_unit_doc_name lostfocus"
''MsgBox txt_unit_doc_name
''
''  d_u_name_Click
'
'
'
'
'End If
'to get doc code over


'***************************
 
DBGrid1.Columns(5).Locked = True
Text22 = " PRESS CONTROL KEY TO COME OUT OF DBGRID .... and TO REFRESH TOTAL AMOUNT"
If DateDiff("d", Format(rs2("consult_dt"), "DD/MM/YYYY"), me_date) > 90 Then
     
     'I = MsgBox(DateDiff("d", Format(rs2("consult_dt")), me_date), , "NO OF DAYS : ")
     'MsgBox "SELECT 'B003' CODE, FOR RECONSULTATION CHARGES...."
     MsgBox "SELECT 'B003' CODE, FOR RECONSULTATION CHARGES...., if necessary."
End If
'DBGrid1.COL = 0
'DBGrid1.SetFocus
cbo_reqno.SetFocus
Exit Sub
MSG:
End Sub

Private Sub view_Click()
Delete.Visible = True
'******** copied form cleartxt function
lbl_cno.Caption = " "
M_number = "  "
me_date.Mask = ""
me_date.Text = ""


ReDim UserData(0 To 5, 0 To 12)
DBGrid1.Refresh
flag = 2
'**********************
cleartxt
'Set rs4 = rdb.OpenResultset("select * from TA107 where io_memo = 'O'", rdOpenKeyset, 3)
'''Set rs4 = rdb.OpenResultset("select * from TA107 where 1 = 2", rdOpenKeyset, 3) '''261099 commented on 171199
'''If rs4.RowCount = 0 Then
 ''' MsgBox "NO MEMOS FOUND ...."
 ''' New1.SetFocus
'''End If

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
txt_unit_doc_name.Enabled = False ''****5599
'**********
'com_mno.Font = "ARIAL"
'com_mno.SetFocus
'M_number.Mask = "M-######"
M_number.Mask = ">?#&######"
'me_cno.Mask = ">??/#####"
me_cno.Text = "__/_____"
M_number.SelStart = 0
M_number.SetFocus
viewflag = True
End Sub
'ADDED ON 9499 FOR REFRESHING SCREEN AFTER EACH OPERATION
Private Sub Refresh1()
'CLEARTXT
 If M_number.Visible = True Then '4599
    M_number = " "
 End If '4599
 If lbl_cno.Visible = True Then '4599
 lbl_cno = " "
 End If '4599
 
 ReDim UserData(0 To 5, 0 To 12)
 DBGrid1.Refresh
 
 'New1.SetFocus
End Sub

Sub VALIDATEAMT()
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
Public Sub BLOODS()
      ' BLOODCOMBO.ZOrder
        DoEvents
       BLOODCOMBO.Left = 500
       BLOODCOMBO.Top = 2000
       BLOODCOMBO.show
       DoEvents
       BLOODCOMBO.BLOODDOC.SetFocus
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
 me_date.Enabled = False
 DBGrid1.Enabled = False
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
zta1071(0) = me_mno.FormattedText
zrsta1071.Requery
If zrsta1071.RowCount > 0 Then
zrsta1071.MoveFirst
DoEvents
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
                   
                   DoEvents
                   If zmatched = False Then
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
                
                DoEvents
                If zmatched = False Then
                   zrsmaptable.MoveFirst
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

Public Sub PATDEPOSIT()
Dim DWT_AMT As Long
        TA116Qr(0) = CASENO
        Set rs_tr = TA116Qr.OpenResultset(2, 3)
        ''''Set rs_tr = rdb.OpenResultset("select * from ta113 where case_no = '" & CASENO & "'", 2, 3)
                    If Len(Trim(rs_tr("SER_AMT"))) = 0 Then
                        SR_AMT = 0
                    Else
                        SR_AMT = rs_tr("ser_amt")
                    End If
                    
     'CHECKING FOR THE WAITLISTING ENTRY IN TA104 starts
     'FOR THE Case Number
      STAT = ""
      WT_AMT = 0
      Call WAIT_STAT(STAT, CASENO, DWT_AMT)
      
     'CHECKING FOR THE WAITLISTING ENTRY IN TA104 ends
       
       
       If STAT = "0" Then
        If DWT_AMT > ((rs_tr("NET_AMT")) - (SR_AMT + lbl_tot_amt)) Then
          MsgBox "Patient is waitlisted in Waiting list , and his Balance amount exceeds the Wailisting Advance Amount ", vbCritical
          STAT = "0"
          Exit Sub
         Else
          STAT = "1"
        End If
       End If
     
                    
                    
                    
                    lbl_tot_amt.Refresh
                     iRow = 0
             If UserData(0, iRow) <> "Z100" Then
                If (SR_AMT + lbl_tot_amt) > (rs_tr("NET_AMT")) Then
                    MsgBox "MEMO AMOUNT EXCEEDS  MAXIMUM AMOUNT IN PATIENT'S DEPOSIT ACCOUNT ; MEMO CANNOT BE ISSUED "
                    ztotser = rs_tr("ser_amt") + CDbl(lbl_tot_amt.Caption)
                    MsgBox "PATIENT'S DEPOSIT BALANCE AMOUNT = " & (rs_tr("NET_AMT") - (ztotser))
                    
                   '***************************
                   '20/02/2002-Parag
                   'Even if the amount exceeds 90% of the total deposited
                   'amount but the balance after reducing the new memo amt
                   'is more than Rs100 the memo can be saved.
                   lbl_tot_amt.Refresh
                   If (rs_tr("NET_AMT") - (ztotser)) >= 100 Then
                       PATDEPO = True
                       MsgBox "MEMO IS TO BE SAVED"
                   Else
                    '----------------------------------
                    'Even if 90% amount is exceeded for Group 3 (Accounts Group)
                    'for users listed in matchup function can save the memo
                    'on the basis of a sanction given by Authorised personnel.
                    '06/08/2001-Parag
                     zvalidid = MatchUp(txt_doc_sec_name.Text)
                     If IsNull(zvalidid) Then
                       PATDEPO = False
                       btn_vis (0)
                       Exit Sub
                     Else
                     PATDEPO = True
                      zconfirm = MsgBox("GET PERMISSION OF HIGHER OFFICIALS TO ISSUE MEMO", vbYesNo + vbDefaultButton2, "COMMIT BOX")
                       If zconfirm = vbYes Then
                          PATDEPO = True
                       Else
                          PATDEPO = False
                          btn_vis (0)
                          Exit Sub
                       End If
                    End If
                  End If 'for Rs100  condition
                '----------------------------------
                Else
                PATDEPO = True
                MsgBox "MEMO IS TO BE SAVED"
                End If
              Else
                If (ztotser) > (rs_tr("NET_AMT")) Then
                    MsgBox " Service amount exceeds balance amount", vbInformation, "Amount Exceeds"
                    PATDEPO = False
                    btn_vis (0)
                    Exit Sub
                Else
                    PATDEPO = True
                    MsgBox "MEMO IS TO BE SAVED"
                End If
              
            End If
    'End If


End Sub


Private Function updatesmart() As Integer
    Dim rdscard As rdoQuery
    Dim num As Integer
    Set rdscard = rdb.CreateQuery("", "update patoolslib.scardacc set ser_amt=ser_amt+?,bal_amt=bal_amt-? where case_no=? and csnr=? AND AIFLAG='A'")
    rdscard(0) = Val(lbl_tot_amt.Caption)
    rdscard(1) = Val(lbl_tot_amt.Caption)
    rdscard(2) = me_cno.FormattedText
    rdscard(3) = Trim(lbl_csnr.Caption)
    rdscard.Execute
    num = rdscard.RowsAffected
    lbl_deposit.Caption = Val(lbl_deposit.Caption) - Val(lbl_tot_amt.Caption)
    updatesmart = num
    
End Function


Private Sub LOCKACCOUNT(CSNR As String)
    Dim rdscard As rdoQuery
    Set rdscard = rdb.CreateQuery("", "update patoolslib.scardacc set aiflag='L' where case_no='" & Trim(me_cno.FormattedText) & "' AND csnr='" & CSNR & "'")
    rdscard.Execute
    rdscard.Close
    Set rdscard = Nothing
End Sub

Public Sub themprint()
CR1.ReportFileName = "C:\pb121298\opth.rpt"
'If flag = 1 Then
CR1.SelectionFormula = "{TA107.memo_no} = '" & me_mno.Text & "'"
'ElseIf flag = 2 Then
'CR1.SelectionFormula = "{TA107.memo_no} = '" & M_number.Text & "'"
'End If

'CR1.Formulas(0) = "accnum = '" & Trim(Right(me_cno.FormattedText, 5) & Left(me_cno.FormattedText, 2)) & "'"
CR1.Formulas(1) = "bal='" & Trim(lbl_deposit.Caption) & "'"
CR1.Destination = crptToPrinter

CR1.Action = 1

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

Public Function GETSMARTCARD(CASENO As String, CSNR As String) As Integer
    Dim RQ_SMARTCARD As rdoQuery
    Dim RS_SMARTCARD As rdoResultset
    
    Set RQ_SMARTCARD = rdb.CreateQuery("S5M", "SELECT * FROM PATOOLSLIB.SCARDACC WHERE CASE_NO=? AND CSNR=?  and aiflag='A'")
    RQ_SMARTCARD(0) = CASENO
    RQ_SMARTCARD(1) = CSNR
    
    Set RS_SMARTCARD = RQ_SMARTCARD.OpenResultset(1, 1)
    
    If RS_SMARTCARD.RowCount > 0 Then
        GETSMARTCARD = 1
        scard_bal_amt = RS_SMARTCARD("bal_amt")
        DisplayPole ("Case No. " & CASENO & "   " & "Balance " & PadSpace(12, RS_SMARTCARD("bal_amt")))
    Else
        GETSMARTCARD = 0
    End If
    
    
    'MsgBox ("Please place the card on the Reader")
    RQ_SMARTCARD.Close
    Set RQ_SMARTCARD = Nothing
    Set RS_SMARTCARD = Nothing
End Function
Public Function PadSpace(num As Integer, str As String) As String
    Do While Not Len(str) = 12
        str = str & " "
        
    Loop
    PadSpace = str
    
    
End Function
Public Sub DisplayPole(MSG As String)
'    MSComm1.RThreshold = 1
'
'    ' When Inputting Data, Input 1 Byte at a time
'    MSComm1.InputLen = 1
'
'    ' 9600 Baud, No Parity, 8 Data Bits, 1 Stop Bit
'    MSComm1.Settings = "9600,N,8,1"
'    ' Disable DTR
'    MSComm1.DTREnable = False
'
'    ' Open COM1
'    MSComm1.CommPort = 1 'im choosing to open port 1 on my pc, u can chage it
'    If MSComm1.PortOpen = True Then
'        MSComm1.PortOpen = False
'    Else
'        MSComm1.PortOpen = True
'    End If
'
'
'    MSComm1.Output = MSG
'
'MSComm1.PortOpen = False
End Sub

Public Sub UpdateTransaction()
    Dim strUpd As String
    Set PCARD = New PISWriteCard.ClsWriteCard
    PCARD.Initialise
    PCARD.ConnectToCard
    PINFO.LastTransDate = me_date.FormattedText
    PINFO.NewTrnsAmt = lbl_deposit.Caption
    PINFO.FinalAmt = lbl_tot_amt.Caption
    PINFO.TrnsId = me_mno.FormattedText
    PINFO.AccountFlag = "A"
    PINFO.LocationCode = "01"
    
'    pinfo.LastTransAmt = "400"
'    pinfo.NewTrnsAmt = 400
'    pinfo.TrnsId = me_mno.FormattedText
'    pinfo.NewTrnsDate = me_date.FormattedText
'    pinfo.NewTrnsAmt = lbl_tot_amt.Caption * 100
    
    strUpd = PCARD.UpdateTransaction(Strkey, PINFO)
    
    '''ADDED FROM HERE FOR SMARTCARD BALANCE AMOUNT AND SYSTEM BALANCE AMOUN
    Dim str As String
    str = PCARD.GetLastTransaction(Strkey, PINFO)
        
    'INSERTERROR
    If Val(PINFO.FinalAmt) <> Val(lbl_deposit.Caption) Then
        Call INSERTERROR(me_mno.FormattedText, me_cno.FormattedText, PINFO.FinalAmt, lbl_deposit.Caption)
    End If
    
    Call SendSMS(Trim(me_cno.FormattedText), Trim(lbl_tot_amt.Caption), Trim(lbl_deposit.Caption), Trim(me_mno.FormattedText))
        'DisplayPole (" Thank You  " & "        " & "Tot Balance " & PadSpace(8, Trim(lbl_deposit.Caption))) ''ADDED ON 14/09/2011 FOR DISPLAYING BALANCE ON POLE

    ''''UPTO HERE
    Set PCARD = Nothing
End Sub
Public Function ValidateCard(CASENO As String) As String
    Dim RQ_SMARTCARD As rdoQuery
    Dim RS_SMARTCARD As rdoResultset
    
    Set RQ_SMARTCARD = rdb.CreateQuery("SMT", "SELECT csnr FROM PATOOLSLIB.SCARDACC WHERE CASE_NO=?  and aiflag='A'")
    RQ_SMARTCARD(0) = CASENO
    
    Set RS_SMARTCARD = RQ_SMARTCARD.OpenResultset(1, 1)
    
    If RS_SMARTCARD.RowCount > 0 Then
        
        ValidateCard = RS_SMARTCARD("csnr")
    Else
        ValidateCard = ""
    End If
    'MsgBox ("Please place the card on the Reader")
    RQ_SMARTCARD.Close
    Set RQ_SMARTCARD = Nothing
    Set RS_SMARTCARD = Nothing
End Function

Public Sub INSERTERROR(MEMONO As String, CASENO As String, CARDBAL As String, SYSBAL As String)
    Dim RQ_INSERTERROR As rdoQuery
    Dim D1 As String
    D1 = Format(me_date.Text, "dd/mm/yyyy")
    D1 = GET_REVDATE(D1)
    
    Set RQ_INSERTERROR = rdb.CreateQuery("INS", "INSERT INTO PATOOLSLIB.SERRORLOG(MEMO_NO,CASE_NO,CARDBALAMT,SYSBALAMT,USERID,IPADDRESS) VALUES(?,?,?,?,?,?)")
    RQ_INSERTERROR(0) = MEMONO
    RQ_INSERTERROR(1) = CASENO
    RQ_INSERTERROR(2) = Val(CARDBAL)
    RQ_INSERTERROR(3) = SYSBAL
    RQ_INSERTERROR(4) = usr_name
'    If Len(Trim(txt_doc_sec_name)) = 0 Then
'        RQ_INSERTERROR(4) = "userid"
'    Else
'        RQ_INSERTERROR(4) = txt_doc_sec_name
'    End If
    RQ_INSERTERROR(5) = ipadd
    RQ_INSERTERROR.Execute
    
    
    Set RQ_INSERTERROR = Nothing
    
End Sub
Private Sub SendSMS(CASENO As String, SERAMT As String, BALANCE As String, MEMONO As String)
    Dim rq_ins As rdoQuery
    Dim MSG As String
    Dim RS_TBLSMSMSGS As rdoResultset
    '''added on 1-1-2013 for serverip
    Dim rs_server As rdoResultset
    Set rs_server = rdb.OpenResultset("select server from patoolslib.serverip where APPLICATN ='SMS'", 1, 1)
    '''upto here
    Dim rs_smsdtls As rdoResultset
    Set rs_smsdtls = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.VWSMSCASES WHERE CASENO='" & CASENO & "' ", 1, 1)
    If Not rs_smsdtls.EOF Then
        If Trim(rs_smsdtls("pmob")) <> "" Then
            
                Set RS_TBLSMSMSGS = rdb.OpenResultset("SELECT * FROM PATOOLSLIB.TBLSMSMSGS WHERE MSGID='12'")
                MSG = RS_TBLSMSMSGS("MSGDESC")
                MSG = Replace(MSG, "@1", SERAMT)
                MSG = Replace(MSG, "@2", CASENO)
                'msg = Replace(msg, "@3", "Services")
                If Val(BALANCE) < 0 Then
                MSG = Replace(MSG, "@3", BALANCE & " DR")
                Else
                MSG = Replace(MSG, "@3", BALANCE & " CR")
                End If
                MSG = Replace(MSG, "@4", MEMONO)
                
                Dim dbsql As New ADODB.Connection
                Dim str As String
                dbsql.Open ("Provider=SQLOLEDB.1;Password=tmhit123;Persist Security Info=True;User ID=SA;Initial Catalog=SMS_DB;Data Source=" & Trim(rs_server(0)))
                'Set rq_ins = dbsql.CreateQuery("", "insert into SMS_IO_LOG_MASTER(SMS_MOBILE,SMS_TEXT,STAMP,PDU_SMS,PRIORITY_INDEX,SMS_METHOD,SMS_IO_IND) values ('" & Trim(rs_smsdtls("MOBILE1")) & "','" & Trim(msg) & "',getdate(),0,'" & Trim(rs_smsdtls("priority")) & "','2','3')")
                str = "insert into SMS_IO_LOG_MASTER(SMS_MOBILE,SMS_TEXT,STAMP,PDU_SMS,PRIORITY_INDEX,SMS_METHOD,SMS_IO_IND,transaction_id,case_no,msgID) values ('" & Trim(rs_smsdtls("pmob")) & "','" & Trim(MSG) & "',getdate(),0,'" & Trim(RS_TBLSMSMSGS("priority")) & "','2','3','" & MEMONO & "','" & CASENO & "','12')"
                dbsql.Execute str
                dbsql.Close
                
           End If
            
'        Set rq_ins = dbsql.CreateQuery("", "insert into SMS_IO_LOG_MASTER(SMS_MOBILE,SMS_TEXT,STAMP,PDU_SMS,PRIORITY_INDEX,SMS_METHOD,SMS_IO_IND) values ('" & Trim(rs_smsdtls("pmob")) & "','" & Trim(msg) & "',getdate(),0,'" & Trim(RS_TBLSMSMSGS("priority")) & "','2','3')")
'
'        rq_ins.Execute
'        rq_ins.CLOSE
'    Set rq_ins = Nothing
    End If
End Sub
