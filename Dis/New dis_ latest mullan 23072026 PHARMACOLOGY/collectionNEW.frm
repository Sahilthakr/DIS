VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form collectionNEW 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form2"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12045
   LinkTopic       =   "Form2"
   ScaleHeight     =   9720
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   11775
      Begin VB.ComboBox cmb_location 
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
         ItemData        =   "collectionNEW.frx":0000
         Left            =   1920
         List            =   "collectionNEW.frx":0002
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin VB.OptionButton OPT_ALL 
         BackColor       =   &H00C0C0C0&
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
         Height          =   315
         Left            =   6240
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton OPT_CASENO 
         BackColor       =   &H00C0C0C0&
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
         Left            =   5040
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1575
         Begin VB.OptionButton Opt_patient 
            BackColor       =   &H00E0E0E0&
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
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opt_staff 
            BackColor       =   &H00E0E0E0&
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
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Width           =   1095
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            FillColor       =   &H00E0E0E0&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   1335
         End
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
         Left            =   840
         TabIndex        =   24
         Top             =   1680
         Width           =   7215
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
         Left            =   9720
         TabIndex        =   13
         Top             =   240
         Width           =   1215
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
         Left            =   7680
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_caseno 
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
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
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
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
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lbl_fromdt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   2520
         TabIndex        =   34
         Top             =   2040
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   120
         TabIndex        =   33
         Top             =   2040
         Width           =   885
      End
      Begin VB.Label lbl_location 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Collected at"
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
         TabIndex        =   32
         Top             =   240
         Width           =   1710
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   240
         TabIndex        =   31
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   8880
         TabIndex        =   30
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   6600
         TabIndex        =   29
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Case No"
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
         TabIndex        =   28
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   11775
      Begin VB.CheckBox chk_collected 
         BackColor       =   &H8000000E&
         Caption         =   "Print Borcode"
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
         Left            =   4800
         TabIndex        =   21
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txt_rowcount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         Left            =   720
         TabIndex        =   22
         Top             =   3120
         Visible         =   0   'False
         Width           =   9375
      End
      Begin VB.CommandButton cmd_print1 
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
         Left            =   120
         TabIndex        =   20
         Top             =   4560
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3375
         Left            =   -120
         TabIndex        =   12
         Top             =   360
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   5953
         _Version        =   393216
         Rows            =   3
         Cols            =   5
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         ForeColorFixed  =   4210816
         BackColorBkg    =   14737632
         GridColor       =   16777215
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"collectionNEW.frx":0004
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
   Begin VB.Frame buttons 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   7920
      Width           =   11775
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "B&arcodePrint"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         MouseIcon       =   "collectionNEW.frx":00A6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton CMD_REFRESH 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&REFRESH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6630
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmd_new 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&NEW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         MaskColor       =   &H00FFC0C0&
         MouseIcon       =   "collectionNEW.frx":03B0
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmd_view 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&VIEW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2010
         MouseIcon       =   "collectionNEW.frx":06BA
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmd_cancel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&CANCEL"
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
         Height          =   495
         Left            =   3555
         MouseIcon       =   "collectionNEW.frx":09C4
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmd_back 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&BACK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5085
         MouseIcon       =   "collectionNEW.frx":0CCE
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmd_scanner 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&SCAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         MaskColor       =   &H00FFC0C0&
         MouseIcon       =   "collectionNEW.frx":0FD8
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   1455
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
      Left            =   10680
      TabIndex        =   16
      Top             =   720
      Width           =   1215
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
      Left            =   10680
      TabIndex        =   15
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   960
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1080
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lbl_req 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRINT  BARCODE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   4320
      TabIndex        =   39
      Top             =   720
      Width           =   2490
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   10080
      TabIndex        =   38
      Top             =   720
      Width           =   420
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   10080
      TabIndex        =   37
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   330
      Left            =   1890
      TabIndex        =   36
      Top             =   240
      Width           =   8595
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Last modified on 02/11/21"
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "collectionNEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''''''''''
''''   Designed and Coded by KANCHAN for printing BARCODE

Dim mydate As String
Dim dept As String
Dim MODCODE As String
Dim modecode As String
Dim rlocation_qr As rdoQuery
Dim rlocation_rs As rdoResultset
Dim rlocationflag As Boolean
Dim rlocation_qrall As rdoQuery
Dim rlocation_rsall As rdoResultset
Dim rlocationflagall As Boolean

Dim P_LAST_WORKING_DAY As String


'Added by Milind on 22-02-2006
Dim selrreqn_opqrall As rdoQuery
Dim selrreqn_oprsall As rdoResultset
Dim selrreqnflagall As Boolean

Dim selrreqn_opqr As rdoQuery
Dim selrreqn_oprs As rdoResultset
Dim selrreqnflag As Boolean


Dim SELREQNO_STFQR As rdoQuery
Dim SELREQNO_STFRS As rdoResultset
Dim SELREQNOSTFFLAG As Boolean

Dim SELREQNO_STFQRALL As rdoQuery
Dim SELREQNO_STFRSALL As rdoResultset
Dim SELREQNOSTFFLAGALL As Boolean

Dim SELRREQUISITNALL_STFQR As rdoQuery
Dim SELRREQUISITNALL_STFRS As rdoResultset
Dim SELRREQUISITNALL_STFFLAGALL As Boolean

Dim SELRREQUISITN_STFQR As rdoQuery
Dim SELRREQUISITN_STFRS As rdoResultset
Dim SELRREQUISITN_STFFLAG As Boolean
'UPTO HERE

Dim COLLECTIONFLAG As Boolean
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG As Boolean

Dim UPDRTIMECN_QR As rdoQuery

Dim SELRREQUISITN_QRVIEW As rdoQuery
Dim SELRREQUISITN_RSVIEW As rdoResultset
Dim SELRREQUISITNFLAGVIEW As Boolean

Dim SELRREQUISITN_QRALL As rdoQuery
Dim SELRREQUISITN_RSALL As rdoResultset
Dim SELRREQUISITNFLAGALL As Boolean

Dim SELRREQUISITN_QRALLVIEW As rdoQuery
Dim SELRREQUISITN_RSALLVIEW As rdoResultset
Dim SELRREQUISITNFLAGALLVIEW As Boolean

Dim updrrequisitn_qr As rdoQuery
Dim SELMA101_QR As rdoQuery
Dim SELMA101_RS As rdoResultset
Dim SELMA101FLAG As Boolean
Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim allflag As Boolean
Dim casenoflag As Boolean
Dim SELREQNO_QR As rdoQuery
Dim SELREQNO_RS As rdoResultset
Dim SELREQNOFLAG As Boolean

Dim SELREQNO_QRVIEW As rdoQuery
Dim SELREQNO_RSVIEW As rdoResultset
Dim SELREQNOFLAGVIEW As Boolean
Dim selta101_qr1 As rdoQuery
Dim selta101_rs1 As rdoResultset
Dim selta101flag1 As Boolean
Dim SELDREQSTAFFD_QR As rdoQuery
Dim SELDREQSTAFFD_RS As rdoResultset
Dim SELDREQSTAFFDFLAG As Boolean
Dim SELRMODLAITYQR As rdoQuery
Dim SELRMODLAITYRS As rdoResultset
Dim SELRMODLAITYFLAG As Boolean

'Dim SELMA106_QR As rdoQuery
'Dim SELMA106_RS As rdoResultset
Dim SELMA106FLAG As Boolean

Dim SELDMULPROCEQR As rdoQuery
Dim SELDMULPROCERS As rdoResultset
Dim SELDMULPROCEFLAG As Boolean
Dim upddmulproceqr As rdoQuery
Dim SELDMULPROCEQRDT As rdoQuery
Dim SELDMULPROCERSDT As rdoResultset
Dim SELDMULPROCEDTFLAG As Boolean

''bloodprodu
Dim SELBLOODPRODU_QR As rdoQuery
Dim SELBLOODPRODU_RS As rdoResultset
Dim SELBLOODPRODU_FLAG As Boolean

Dim SELRREQUISITN_QRCASE As rdoQuery
Dim SELRREQUISITN_RSCASE As rdoResultset
Dim SELRREQUISITN_FLAGCASE As Boolean

Dim LAST_W_QR As rdoQuery
Dim LAST_W_RS As rdoResultset
Dim LAST_W_FLG As Boolean

''**ADD ON 30/10/03 BY MONO
Dim SELPATHNO_QR As rdoQuery
Dim SELPATHNO_RS As rdoResultset
Dim SELPATHNO_FLAG As Boolean
Dim SELALLMODALITY_QRVIEW As rdoQuery
Dim SELALLMODALITY_RSVIEW As rdoResultset
Dim SELALLMODALITY_FLAGVIEW As Boolean

Dim SELRREQUISITNALL_QR As rdoQuery
Dim SELRREQUISITNALL_RS As rdoResultset
Dim SELRREQUISITNALL_FLAG As Boolean

'''''barcode
Dim seldmulproce_qr1 As rdoQuery
Dim seldelprosub_qr As rdoQuery
Dim SAMPLECODE_QR1 As rdoQuery

Dim SELPROCODE_QR As rdoQuery

Dim SEL_DMUL_QR As rdoQuery
Dim SEL_DMUL_RS As rdoResultset

Dim SEL_MA101_QR As rdoQuery
Dim SEL_MA101_RS As rdoResultset

Dim TESTDESC As String

Dim SELLABNO_QR As rdoQuery
Dim SELLABNO_RS As rdoResultset
Dim SELLABNOSF_QR As rdoQuery
Dim SELLABNOSF_RS As rdoResultset
Dim BARCODE_NEW As Boolean
Dim SELTIME_QR As rdoQuery
Dim SELTIME_RS As rdoResultset
Dim SELPATH_QR As rdoQuery
Dim SELPATH_RS As rdoResultset
Dim SELPROCODE_QR1 As rdoQuery
Dim SELPROCODE_RS1 As rdoResultset

Dim DBCPROSUB_QR As rdoQuery
Dim DBCPROSUB_RS As rdoResultset
Dim INS_BARCODE As rdoQuery
Dim newbarcode As Boolean

Dim SELDEPTDES_QR As rdoQuery
Dim SELDEPTDES_RS As rdoResultset
Dim VIEWFLAG_BCD As Boolean
Dim NEWFLAG_BCD As Boolean
Dim colorcd As String
Dim smplcd As String
Dim DESCRIP As String


Dim REQ22 As Variant
Public row1, reqvw
'Dim cstaff As Integer
Dim c As Integer
'***** Shraddha Pandit 28/05/2009 ******
Dim sel_barlbl_qr As rdoQuery
Dim sel_barlbl_rs As rdoResultset
Dim VF As Boolean
'***************************************
'****** Shraddha Pandit 16/07/09 *******
Dim sel_ssrreqn_qr As rdoQuery
Dim sel_ssrreqn_rs As rdoResultset
'***************************************
'shweta 07-feb-2012 for barcode patnstat
Dim SELRREQUISITNPSTAT_QR As rdoQuery
Dim SELRREQUISITNPSTAT_RS As rdoResultset
''***********************

Private Sub chk_collected_Click()
'If chk_collected.Value = 1 Then
'    MSFlexGrid1.Text = "Print Barcode"
'Else
'    MSFlexGrid1.Text = ""
'End If
'** Commented and added by Shraddha 06/07/09 **
If chk_collected.Value = 1 Then
    MSFlexGrid1.Text = "Print Barcode"
Else
    MSFlexGrid1.Text = ""
End If
'**********************************************
End Sub
Private Sub chk_collected_KeyPress(KeyAscii As Integer)
On Error GoTo MSG:
If KeyAscii = 13 Then
    If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
    End If
    If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.row = MSFlexGrid1.row + 1
        chk_collected = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))
        chk_collected.SetFocus
        chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    End If
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmb_dept_Click()
Dim i As Integer
If Trim(cmb_dept.Text) = "ALL" Then
    cmb_modality.Enabled = False
    'cmb_location.Enabled = False
    'Opt_all.Enabled = False
    'msk_caseno.Enabled = False
    OPT_CASENO.Enabled = False
    msk_todt.Enabled = True
    msk_fromdt.Enabled = True
    'msk_fromdt.Text = DATE11
    msk_fromdt.Text = P_LAST_WORKING_DAY
    msk_todt.Text = DATE11
    msk_fromdt.Text = mydate
       ' msk_fromdt.SetFocus
'by albert for modification shifted to form load***************
'    If rlocationflagall = False Then
'        Set rlocation_rsall = rlocation_qrall.OpenResultset(1, 3)
'        rlocationflagall = True
'   Else
'        rlocation_rsall.MoreResults
'        rlocation_rsall.Requery
'    End If
'    If rlocation_rsall.RowCount > 0 Then
'        cmb_location.clear
'        While Not rlocation_rsall.EOF
'            cmb_location.AddItem Trim(rlocation_rsall("LOCATION"))
'            rlocation_rsall.MoveNext
'        Wend
'    End If
'''''*****************
Else
i = InStr(cmb_dept, "$")
dept = Mid(cmb_dept, i + 1, 4)

If SELRMODLAITYFLAG = False Then
    SELRMODLAITYQR(0) = Trim(dept)
    Set SELRMODLAITYRS = SELRMODLAITYQR.OpenResultset(1, 3)
    SELRMODLAITYFLAG = True
Else
    SELRMODLAITYQR(0) = Trim(dept)
    SELRMODLAITYRS.MoreResults
    SELRMODLAITYRS.Requery
End If
  If SELRMODLAITYRS.RowCount > 0 Then
    cmb_modality.clear
    While Not SELRMODLAITYRS.EOF
        cmb_modality.AddItem Trim(SELRMODLAITYRS("mODLITYDES")) & "$" & Trim(SELRMODLAITYRS("MODALITYCD"))
        SELRMODLAITYRS.MoveNext
    Wend
  End If
End If
End Sub
Private Sub cmb_location_Click()
'msk_fromdt.Text = DATE11
'msk_todt.Text = DATE11
'msk_fromdt.SetFocus
''If allflag = True Then
''        OPT_ALL.Value = True
''    End If
''    If casenoflag = True Then
''        OPT_CASENO.Value = True
''        Msk_caseno.Enabled = True
''    End If
End Sub

Private Sub cmb_location_LostFocus()
OPT_CASENO.Value = True
'OPT_CASENO.SetFocus

End Sub

Private Sub cmb_modality_Click()
'' by albert
Dim i As Integer
i = InStr(cmb_modality, "$")
modecode = Mid(cmb_modality, i + 1, 2)
Exit Sub
''/****** upto
'Dim i As Integer
i = InStr(cmb_modality, "$")
modecode = Mid(cmb_modality, i + 1, 2)

    If rlocationflag = False Then
        rlocation_qr(0) = Trim(dept)
        rlocation_qr(1) = Trim(modecode)
        Set rlocation_rs = rlocation_qr.OpenResultset(1, 3)
        rlocationflag = True

   Else
        rlocation_qr(0) = Trim(dept)
        rlocation_qr(1) = Trim(modecode)
        rlocation_rs.MoreResults
        rlocation_rs.Requery
    End If
    If rlocation_rs.RowCount > 0 Then
        cmb_location.clear
        While Not rlocation_rs.EOF
            cmb_location.AddItem Trim(rlocation_rs("LOCATION"))
            rlocation_rs.MoveNext
        Wend
    End If
'If MODECODE = "BC" Or MODECODE = "EL" Then
'    S.No |^    Requisation No.    |^    Case No.    |^   Procedure Name   |^   Payment     |^  Sample Collected
'End If

End Sub

Private Sub cmd_back_Click()
    
BARCODE_NEW = False
    cmb_location.ListIndex = -1
    cmb_location.Enabled = False
    msk_fromdt.Enabled = False
    msk_todt.Enabled = True
    chk_collected.Visible = False
    msk_fromdt.Mask = ""
    msk_fromdt.Text = ""
    msk_fromdt.Mask = "##/##/####"
    msk_todt.Mask = ""
    msk_todt.Text = ""
    msk_todt.Mask = "##/##/####"
    txt_name.Text = ""
    OPT_ALL.Value = False
    OPT_CASENO.Value = False
    msk_caseno.Enabled = False
    Opt_patient.Enabled = False
    opt_staff.Enabled = False
    VF = False '28/05/2009
    chk_collected.Value = 0
    MSFlexGrid1.Rows = 1
btnenabl ("INI")
Me.MousePointer = 11
Me.Hide
DIS_MAIN.show
'Unload Me
Me.MousePointer = 0
End Sub

Private Sub cmd_cancel_Click()
VF = False '28/05/2009
Image1.Picture = LoadPicture("")
'cmd_scanner.Enabled = False
cmb_dept.clear
'cmb_location.clear
cmb_modality.clear
msk_fromdt.Text = ""
'Msk_todt.Text = ""
msk_todt.Text = ""
txt_rowcount.Text = ""

BARCODE_NEW = False
newbarcode = False

MSFlexGrid1.Rows = 1
If OPT_ALL.Enabled = True Then
    cmb_location.Enabled = False
    msk_fromdt.Enabled = False
    msk_todt.Enabled = True
    chk_collected.Visible = False
    msk_fromdt.Mask = ""
    msk_fromdt.Text = ""
    msk_fromdt.Mask = "##/##/####"
    msk_todt.Mask = ""
    msk_todt.Text = ""
    msk_todt.Mask = "##/##/####"
    txt_name.Text = ""
    OPT_ALL.Value = False
    OPT_CASENO.Value = False
    msk_caseno.Enabled = False
    Opt_patient.Enabled = False
    opt_staff.Enabled = False
    cmb_modality.Enabled = False
    cmb_dept.Enabled = False
    cmd_print1.Visible = False
    allflag = True
End If
If OPT_CASENO.Enabled = True Then
    cmb_location.Enabled = False
    msk_fromdt.Enabled = False
    msk_todt.Enabled = True
    chk_collected.Visible = False
    msk_fromdt.Mask = ""
    msk_fromdt.Text = ""
    txt_name.Text = ""
    msk_fromdt.Mask = "##/##/####"
    msk_todt.Mask = ""
    msk_todt.Text = ""
    msk_todt.Mask = "##/##/####"
    OPT_ALL.Value = False
    OPT_CASENO.Value = False
    msk_caseno.Enabled = False
    Opt_patient.Enabled = False
    opt_staff.Enabled = False
    msk_caseno.Mask = ""
    msk_caseno.Text = ""
    cmb_modality.Enabled = False
    msk_caseno.Enabled = False
    cmb_dept.Enabled = False
    casenoflag = True
    cmd_print1.Visible = False
End If

btnenabl ("INI")
NEWFLAG = False
viewflag = True

End Sub

Private Sub cmd_modify_Click()
btnenabl ("NEW")
chk_collected.Value = 0
chk_collected.Visible = True
End Sub

Private Sub cmd_new_Click()
On Error GoTo MSG:

'cmd_scanner.Enabled = True


VF = False '28/05/2009
'cmb_location.Text = "GEN.OPD" 'SHWETA 04-NOV-09
'cmb_location.Text = "PVT.OPD" 'Commented and added by Shraddha on 04/02/10
NEWFLAG = True
viewflag = False
BARCODE_NEW = True
newbarcode = True
btnenabl ("NEW")
    'Opt_patient.Value = True
   '''' by albert  insert//*******************

    ''''up to by albert //*******************

    If SELMA106FLAG = False Then
        Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
        SELMA106FLAG = True
    Else
        SELMA106_RS.MoreResults
        SELMA106_RS.Requery

    End If
    If SELMA106_RS.RowCount > 0 Then
        cmb_dept.clear
        cmb_dept.AddItem "ALL"
        While Not SELMA106_RS.EOF
            cmb_dept.AddItem Trim(SELMA106_RS("dept_name")) & "$" & Trim(SELMA106_RS("dept_code"))
            SELMA106_RS.MoveNext
        Wend
    End If
msk_fromdt.Text = P_LAST_WORKING_DAY
        msk_todt.Text = DATE11
        msk_fromdt.Text = mydate

'    If rlocationflag = False Then
'        Set rlocation_rs = rlocation_qr.OpenResultset(1, 3)
'        rlocationflag = True
'
'   Else
'        rlocation_rs.MoreResults
'        rlocation_rs.Requery
'    End If
'    If rlocation_rs.RowCount > 0 Then
'        cmb_location.clear
'        While Not rlocation_rs.EOF
'            cmb_location.AddItem Trim(rlocation_rs("LOCATION"))
'            rlocation_rs.MoveNext
'        Wend
'    End If
'If SELRMODLAITYFLAG = False Then
'    Set SELRMODLAITYRS = SELRMODLAITYQR.OpenResultset(1, 3)
'    SELRMODLAITYFLAG = True
'Else
'    SELRMODLAITYRS.MoreResults
'    SELRMODLAITYRS.Requery
'End If
'  If SELRMODLAITYRS.RowCount > 0 Then
'    cmb_modality.clear
'    While Not SELRMODLAITYRS.EOF
'        cmb_modality.AddItem Trim(SELRMODLAITYRS("mODLITYDES")) & "$" & Trim(SELRMODLAITYRS("MODALITYCD"))
'        SELRMODLAITYRS.MoveNext
'    Wend
'  End If
    cmb_location.Enabled = True
    txt_name.Text = ""
    msk_fromdt.Enabled = False
    msk_todt.Enabled = True
    OPT_ALL.Enabled = True
    OPT_CASENO.Enabled = True
    Frame3.Visible = False
    Opt_patient.Enabled = True
    opt_staff.Enabled = True
    msk_caseno.Visible = False
    msk_caseno.Enabled = False
    Opt_patient.Value = False
    opt_staff.Value = False
    txt_name.Visible = False
    lbl_name.Visible = False
    Label4.Visible = False
    cmb_dept.Visible = False
    cmb_modality.Visible = False
    Label3.Visible = False
    Label1.Visible = False
    'cmb_dept.Enabled = True
    'cmb_modality.Enabled = True
     cmb_location.SetFocus
   ' cmb_dept.SetFocus
    cmd_print1.Visible = True
'    cmb_location.Text = ""
    OPT_CASENO.Value = True
    
msk_fromdt.Text = P_LAST_WORKING_DAY
Exit Sub
'NEWFLAG = True

MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_print_Click()
Dim var_date_f
Dim var_date_t
Dim str As String
str = "C"
On Error GoTo MSG:
str = ipadd & " " & Format(DATE11, "dd/mm/yyyy") & " " & time1
var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
If OPT_ALL.Value = True And Trim(cmb_dept.Text) = "ALL" Then
    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\COLLECTION1.RPT"
    CrystalReport1.SelectionFormula = "{COLLECTLF.APNTDT} >= Date(" & var_date_f & ") AND {COLLECTLF.APNTDT} <= Date(" & var_date_t & ") AND {COLLECTLF.LOCATION}='" & Trim(cmb_location.Text) & "'" ' AND {COLLECTLF.REQNSTAT}='C'" 'AND {RREQUISITN.APNTDT}=DATE(" & Format(DATE11, "YYYY,MM,DD") & ")"
End If
If OPT_ALL.Value = True And Trim(cmb_dept.Text) <> "ALL" And Trim(cmb_modality.Text) = "" Then
    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\COLLECTION1.RPT"
    CrystalReport1.SelectionFormula = "{COLLECTLF.APNTDT} >= Date(" & var_date_f & ") AND {COLLECTLF.APNTDT} <= Date(" & var_date_t & ") AND {COLLECTLF.LOCATION}='" & Trim(cmb_location.Text) & "' AND  {COLLECTLF.DEPT_CODE}='" & Trim(dept) & "'" 'AND {RREQUISITN.APNTDT}=DATE(" & Format(DATE11, "YYYY,MM,DD") & ")"
End If
If OPT_ALL.Value = True And Trim(cmb_dept.Text) <> "ALL" And Trim(cmb_modality.Text) <> "" Then
    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\COLLECTION1.RPT"
    CrystalReport1.SelectionFormula = "{COLLECTLF.APNTDT} >= Date(" & var_date_f & ") AND {COLLECTLF.APNTDT} <= Date(" & var_date_t & ") AND {COLLECTLF.LOCATION}='" & Trim(cmb_location.Text) & "' AND {COLLECTLF.DEPT_CODE}='" & Trim(dept) & "' AND {COLLECTLF.MODALITYCD}='" & Trim(modecode) & "'"
End If
If OPT_CASENO.Value = True Then
    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\COLLECTIONC.RPT"
    CrystalReport1.SelectionFormula = "{RREQUISITN.CASENO}= '" & msk_caseno.FormattedText & "' and {RREQUISITN.APNTDT} >= Date(" & var_date_f & ") AND {RREQUISITN.APNTDT} <= Date(" & var_date_t & ") AND {RREQUISITN.LOCATION}='" & Trim(cmb_location.Text) & "'"

    'AND {RREQUISITN.REQNSTAT}='C'  AND {COLLECTLF.APNTDT}=DATE(" & Format(DATE11, "YYYY,MM,DD") & ")"
End If
'     If Trim(cmb_dept.Text) = "ALL" Then
'        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\COLLECTIONC.RPT"
'        CrystalReport1.SelectionFormula = "{RREQUISITN.CASENO}= '" & msk_caseno.FormattedText & "'"
'        CrystalReport1.Formulas(0) = "LOCATION = '" & cmb_location.Text & "'"
'        CrystalReport1.Formulas(1) = "USRID = '" & STR & "'"
'     End If
CrystalReport1.Formulas(0) = "FROMDT = '" & msk_fromdt.FormattedText & "'"
CrystalReport1.Formulas(1) = "TODT = '" & msk_todt.FormattedText & "'"
CrystalReport1.Formulas(2) = "LOCATION = '" & cmb_location.Text & "'"
CrystalReport1.Formulas(3) = "USRID = '" & str & "'"

CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

        CrystalReport1.Action = 1
btnenabl ("PRINT")
Exit Sub
MSG:
 MsgBox Err.Description
End Sub

Private Sub CMD_PRINT1_Click()
    str = ipadd & " " & Format(DATE11, "dd/mm/yyyy") & " " & time1
    If Trim(cmb_dept.Text) = "ALL" Then
        var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
        var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\COLLECTIONNALL.RPT"
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ")  "



        CrystalReport1.Formulas(0) = "FROMDT = '" & msk_fromdt.FormattedText & "'"
        CrystalReport1.Formulas(1) = "TODT = '" & msk_todt.FormattedText & "'"
        CrystalReport1.Formulas(2) = "LOCATION = '" & cmb_location.Text & "'"
        CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
     End If
     If Trim(cmb_dept.Text) <> "ALL" And OPT_CASENO.Value = True Then
        'var_date_f = Year(msk_fromdt.FormattedText) & "," & Month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
        'var_date_t = Year(msk_todt.FormattedText) & "," & Month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\COLLECTIONCCASE.RPT"
        'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ")  "
        CrystalReport1.SelectionFormula = "{RREQUISITN.CASENO}= '" & msk_caseno.FormattedText & "' AND {RREQUISITN.MODALITYCD} ='" & modecode & "'"


        'CrystalReport1.Formulas(0) = "FROMDT = '" & msk_fromdt.FormattedText & "'"
        'CrystalReport1.Formulas(1) = "TODT = '" & msk_todt.FormattedText & "'"
        'CrystalReport1.Formulas(2) = "LOCATION = '" & cmb_location.Text & "'"
         CrystalReport1.Formulas(3) = "USRID = '" & str & "'"

     End If
     If OPT_ALL.Value = True And Trim(cmb_dept.Text) <> "ALL" Then
        var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
        var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\COLLECTIONCCASE.RPT"
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD} ='" & modecode & "'"
        'CrystalReport1.SelectionFormula = "{RREQUISITN.CASENO}= '" & msk_caseno.FormattedText & "' AND {RREQUISITN.MODALITYCD} ='" & modecode & "'"


        CrystalReport1.Formulas(0) = "FROMDT = '" & msk_fromdt.FormattedText & "'"
        CrystalReport1.Formulas(1) = "TODT = '" & msk_todt.FormattedText & "'"
        'CrystalReport1.Formulas(2) = "LOCATION = '" & cmb_location.Text & "'"
         CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
     End If
     CrystalReport1.WindowState = crptMaximized
     CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

     CrystalReport1.Action = 1
End Sub

Private Sub cmd_refresh_Click()
VF = False '28/05/2009
If Opt_patient.Value = True Then
    msk_caseno.Mask = ""
    msk_caseno.Text = ""
    msk_caseno.Mask = "??>/#####"
ElseIf opt_staff.Value = True Then
    msk_caseno.Mask = ""
    msk_caseno.Text = ""
    msk_caseno.Mask = "#######"
End If

'msk_fromdt.Mask = ""
'msk_fromdt.Text = DATE11
'
'msk_todt.Mask = ""
'msk_todt.Text = DATE11


NEWFLAG = True
MSFlexGrid1.Rows = 1
'txt_pathno.Text = ""
'chk_collected.Value = 0
chk_collected.Visible = False
'txt_pathno.Visible = False
btnenabl ("NEW")
MODIFYFLAG = False
'chk_rcvd.Visible = False
If msk_caseno.Enabled = True Then
    msk_caseno.SetFocus
End If
txt_name.Text = ""
End Sub

Private Sub cmd_save_Click()
 On Error GoTo MSG:
 If Len(Trim(cmb_location.Text)) = 0 Then
   MsgBox " Select Sample collected At....."
 Else
 rdb.BeginTrans
 For i = 1 To MSFlexGrid1.Rows - 1
    MSFlexGrid1.col = 5

    If Trim(MSFlexGrid1.TextMatrix(i, 5)) <> "" Then
        updrrequisitn_qr(0) = Trim(cmb_location.Text)
        updrrequisitn_qr(1) = "C"
        updrrequisitn_qr(2) = GET_REVDATE(txt_date.Text)
        updrrequisitn_qr(3) = req_chglen(MSFlexGrid1.TextMatrix(i, 2))
        updrrequisitn_qr.Execute

    UPDRTIMECN_QR(0) = GET_REVDATE(Trim(txt_date.Text))
    UPDRTIMECN_QR(1) = Format(Trim(txt_time.Text), "HH:MM:SS")
    UPDRTIMECN_QR(2) = req_chglen(MSFlexGrid1.TextMatrix(i, 2))
    UPDRTIMECN_QR.Execute


    Else
        If Trim(MSFlexGrid1.TextMatrix(i, 2)) <> "" And Trim(MSFlexGrid1.TextMatrix(i, 3) <> "") Then
            updrrequisitn_qr(0) = ""
            updrrequisitn_qr(1) = "P"
            updrrequisitn_qr(2) = GET_REVDATE(txt_date.Text)
            updrrequisitn_qr(3) = req_chglen(MSFlexGrid1.TextMatrix(i, 2))
            updrrequisitn_qr.Execute
        End If
    End If

Next i

If OPT_ALL.Value = True Then

End If
'If opt_caseno.Value = True Then
'
'     If SELREQNOFLAG = False Then
'            SELREQNO_QR(0) = msk_caseno.FormattedText
'            SELREQNO_QR(1) = "P"
'            SELREQNO_QR(2) = "P"
'            SELREQNO_QR(3) = Trim(dept)
'            i = InStr(cmb_modality, "$")
'            SELREQNO_QR(4) = Mid(cmb_modality, i + 1, 2)
'            Set SELREQNO_RS = SELREQNO_QR.OpenResultset(1, 3)
'            SELREQNOFLAG = True
'        Else
'           SELREQNO_QR(0) = msk_caseno.FormattedText
'           SELREQNO_QR(1) = "P"
'           SELREQNO_QR(2) = "P"
'           SELREQNO_QR(3) = Trim(dept)
'           i = InStr(cmb_modality, "$")
'           SELREQNO_QR(4) = Mid(cmb_modality, i + 1, 2)
'
'           SELREQNO_RS.MoreResults
'           SELREQNO_RS.Requery
'
'        End If
'
'        If SELREQNO_RS.RowCount > 0 Then
'             i = 1
'                MSFlexGrid1.Rows = 1
'                While Not SELREQNO_RS.EOF
'                If Trim(SELREQNO_RS("ser_code")) <> "8888" Then
'
'                     If SELMA101FLAG = False Then
'                         SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
'                         Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
'                         SELMA101FLAG = True
'                     Else
'                         SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
'                         SELMA101_RS.MoreResults
'                         SELMA101_RS.Requery
'                    End If
''                     If SELMA101_RS.RowCount > 0 Then
''
''                     End If
'                Else
''                     If seldmulproceflag = False Then
''                         seldmulproceqr(0) = msk_caseno.FormattedText
''                         Set seldmulprocers = seldmulproceqr.OpenResultset(1, 3)
''                         seldmulproceflag = True
''                     Else
''                         seldmulproceqr(0) = msk_caseno.FormattedText
''                         seldmulprocers.MoreResults
''                         seldmulprocers.Requery
''                     End If
''                     If seldmulprocers.RowCount > 0 Then
''                        If SELMA101FLAG = False Then
''                         SELMA101_QR(0) = seldmulprocers("pro_CODE")
''                         Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
''                         SELMA101FLAG = True
''                        Else
''                         SELMA101_QR(0) = seldmulprocers("pro_CODE")
''                         SELMA101_RS.MoreResults
''                         SELMA101_RS.Requery
''                        End If
''                        If SELMA101_RS.RowCount > 0 Then
''                            upddmulproceqr(0) = "COLLECTED"
''                            upddmulproceqr(1) = SELDMULPROCERS("pro_CODE")
''                            upddmulproceqr(2) = SELREQNO_RS("REQNNO")
''                            upddmulproceqr.Execute
''
''                        End If
''                     End If
'
'                End If
'                SELREQNO_RS.MoveNext
'                Wend
'               ' chk_collected.Visible = True
'        End If
'
'End If
rdb.CommitTrans
'COMMENTED BY SHWETA AT SANGRUR 26-08-2015
'MsgBox "Record is Successfully Saved."
btnenabl ("SAVE")
End If
Exit Sub
MSG:
    rdb.RollbackTrans
    MsgBox Err.Description

End Sub

Private Sub cmd_scanner_Click()
showmain = True
'***** Shraddha 29/07/2011 *****
dept_collection_BARCODE.cmb_location.Text = Trim(cmb_location.Text)
'*******************************
Me.MousePointer = 11
dept_collection_BARCODE.show
Me.Hide
Me.MousePointer = 0

End Sub

Private Sub cmd_view_Click()
'On Error GoTo MSG:

'cmd_scanner.Enabled = False

VF = True '28/05/2009
btnenabl ("VIEW")
viewflag = True
NEWFLAG = False
newbarcode = False
msk_fromdt.Text = P_LAST_WORKING_DAY
msk_todt.Text = DATE11
msk_fromdt.Text = mydate
cmb_location.Text = ""
''' by albert ///*****
'If rlocationflagall = False Then
'        Set rlocation_rsall = rlocation_qrall.OpenResultset(1, 3)
'        rlocationflagall = True
'   Else
'        rlocation_rsall.MoreResults
'        rlocation_rsall.Requery
'    End If
'    If rlocation_rsall.RowCount > 0 Then
'        cmb_location.clear
'        While Not rlocation_rsall.EOF
'            cmb_location.AddItem Trim(rlocation_rsall("LOCATION"))
'            rlocation_rsall.MoveNext
'        Wend
'    End If
'''' up to albert /////////***************
  If SELMA106FLAG = False Then
        Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
        SELMA106FLAG = True
    Else
        SELMA106_RS.MoreResults
        SELMA106_RS.Requery

    End If
    If SELMA106_RS.RowCount > 0 Then
        cmb_dept.clear
        cmb_dept.AddItem "ALL"
        While Not SELMA106_RS.EOF
            cmb_dept.AddItem Trim(SELMA106_RS("dept_name")) & "$" & Trim(SELMA106_RS("dept_code"))
            SELMA106_RS.MoveNext
        Wend
    End If


'If rlocationflag = False Then
'        Set rlocation_rs = rlocation_qr.OpenResultset(1, 3)
'        rlocationflag = True
'
'   Else
'        rlocation_rs.MoreResults
'        rlocation_rs.Requery
'    End If
'    If rlocation_rs.RowCount > 0 Then
'        cmb_location.clear
'        While Not rlocation_rs.EOF
'            cmb_location.AddItem Trim(rlocation_rs("LOCATION"))
'            rlocation_rs.MoveNext
'        Wend
'    End If
'    If SELRMODLAITYFLAG = False Then
'        Set SELRMODLAITYRS = SELRMODLAITYQR.OpenResultset(1, 3)
'        SELRMODLAITYFLAG = True
'    Else
'        SELRMODLAITYRS.MoreResults
'        SELRMODLAITYRS.Requery
'    End If
'      If SELRMODLAITYRS.RowCount > 0 Then
'        cmb_modality.clear
'        While Not SELRMODLAITYRS.EOF
'            cmb_modality.AddItem Trim(SELRMODLAITYRS("mODLITYDES")) & "$" & Trim(SELRMODLAITYRS("MODALITYCD"))
'            SELRMODLAITYRS.MoveNext
'        Wend
'      End If
    txt_name.Text = ""
chk_collected.Visible = False
cmb_location.Enabled = True
msk_fromdt.Enabled = True
msk_todt.Enabled = True
OPT_ALL.Enabled = True
OPT_CASENO.Enabled = True
Opt_patient.Enabled = True
opt_staff.Enabled = True
Opt_patient.Value = False
opt_staff.Value = False
cmb_modality.Enabled = True
cmb_dept.Enabled = True
cmd_print1.Visible = False
'cmb_location.SetFocus
'cmb_dept.SetFocus
 
    'cmb_location.Text = "GEN.OPD"
    'cmb_location.Text = "PVT.OPD" 'added by Shraddha on 04/02/10

    OPT_CASENO.Value = True
    
    
    msk_fromdt.Text = P_LAST_WORKING_DAY
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Command1_Click()


'''''''''barcode BY Kanchan
Dim bpath As String
'bpath = "C:\disreportsec\data1" 'path commented and added by Bhushan
bpath = "C:\DISREPORT\BARCODE"
Dim QTY1 As String
Dim BAR_CD As String
Dim EXP_DT As String
Dim ITEM_DESC As String
Dim TEMPDATA As String
Dim rack As String
Dim batchno1 As String
Dim itemcode1 As String
Dim i As Long, START
Dim j As Integer
Dim NOL As Integer
Dim Nos As Integer
Dim PRNTSTR As String
Dim PRNTSTR1 As String
Dim row As Integer
Dim req As String
Dim DPT As String
Dim agesex As String
Dim delay As Integer
Dim REQ111t, P_CODE111t, SCODE111t
'SHWETA 30-01-2012
Dim BP1, BP2, BP3 As String
''''ADDED BY CHARULATA ON 14/09/2010 FOR DATETIME
datetime
txt_date.Text = DATE11
txt_time.Text = time1

'''UPTO HERE
'ADDED BY SHWETA ON 07-FEB-2012 FOR PATN STAT TO BARCODE
Dim STRSH As String
Dim txt_status As String
Dim CHK_CASENO As String
Dim CHK_REQNNO As String


''****
''''added by charulata for taking the caseno from rrequisitn on 05-03-2012
Dim rs_reqnno As rdoResultset
Dim tblcaseno As String

'SHWETA 09-03-2012 FOR PAT CATG
Dim PCATGS As String

For i = 1 To MSFlexGrid1.Rows - 1 'Uncommented by Shraddha on 04/02/10

    
    delay = 0
    REQNO = req_chglen(Left(Trim(MSFlexGrid1.TextMatrix(i, 2)), 16))
    
    'ADDED BY SHWETA ON 07-FEB-2012 FOR PATN STAT TO BARCODE
    STRSH = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO = ?"
    Set SELRREQUISITNPSTAT_QR = rdb.CreateQuery("", STRSH)
    SELRREQUISITNPSTAT_QR(0) = REQNO
    Set SELRREQUISITNPSTAT_RS = SELRREQUISITNPSTAT_QR.OpenResultset(1, 3)
    
    If SELRREQUISITNPSTAT_RS.EOF Then
    Else
    
    txt_status = SELRREQUISITNPSTAT_RS("PATNSTAT")
    CHK_CASENO = SELRREQUISITNPSTAT_RS("CASENO")
    CHK_REQNNO = SELRREQUISITNPSTAT_RS("REQNNO")
    PCATGS = SELRREQUISITNPSTAT_RS("PAT_CATG")
    End If
    ''****
    If Trim(CHK_CASENO) <> Trim(MSFlexGrid1.TextMatrix(i, 1)) Then
        MsgBox ("CASE NO. NOT MATCHING.PLEASE CHECK...")
        Exit Sub
    End If
    'reqno = req_chglen(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))
    If MSFlexGrid1.TextMatrix(i, 3) = "" Then
    'Commented and added by Shraddha on 03/02/10
    'If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3) = "" Then

    Else
        '************** 28/05/2009 *******
        'Check if label is already printed then don't allow to take print again
        sel_barlbl_qr(0) = REQNO
        Set sel_barlbl_rs = sel_barlbl_qr.OpenResultset(1, 3)
        If sel_barlbl_rs.EOF Or VF = True Then
            GoTo AA
        End If
        'If sel_barlbl_rs("LABELS") <> 0 Then
        If Not sel_barlbl_rs.EOF <> 0 Then
            'MsgBox "Barcode print already taken for " & REQNO
            GoTo AA
        Else
AA:
            
            If Mid(REQNO, 4, 2) = "HL" Then
                            If Len(msk_caseno) = 6 Then 'Staff
                                f1 = FreeFile()
                                
                                Open bpath & "\micro.TXT" For Output As #f1
                                
                                Dim rs_emp_1 As rdoResultset
                
                                Dim DPT_1 As String
                                
                                str = "SELECT distinct SAMPLECODE.SAMPLECD FROM DISLIB.SAMPLECODE,DISLIB.DMULPROCE WHERE DMULPROCE.PRO_CODE=SAMPLECODE.SERCODE AND DMULPROCE.REQNO=?"
                                Set SELLABNO_QR = rdb.CreateQuery("SELLABNO", str)
                
                                SELLABNO_QR(0) = REQNO
                                SELLABNO_QR.Execute
                
                                c = 0
                                Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
                        
                                c = SELLABNO_RS.RowCount
                                colorcd = ""
                                While Not SELLABNO_RS.EOF
                                    REQ111t = REQNO
                
                                    SCODE111t = SELLABNO_RS(0)
                                    smplcd = SELLABNO_RS(0)
                                    
                                    REQ22 = Mid$(REQ111t, 4)
                                    SCODE111t = REQ22 & SCODE111t
                
                                    deptdesc
                                    
                                    SELLABNO_RS.MoveNext
                                Wend
                                    
                                    DPT_1 = Mid(REQNO, 1, 3)
                                    
                                    SELDEPTDES_QR(0) = DPT_1
                                    Set SELDEPTDES_RS = SELDEPTDES_QR.OpenResultset(1, 3)
                                                            
                                    DESCRIP = ""
                                    DESCRIP = Mid(Trim(SELDEPTDES_RS("DEPT_NAME")), 1, 4) & "/" & colorcd
                                    
                                    Set rs_emp_1 = rdb.OpenResultset("select age,sex,PNAME from dislib.dreqstaffd where empcode='" & Trim(msk_caseno.FormattedText) & "'", 1, 1)
                                    agesex1 = rs_emp_1(0) & "/" & rs_emp_1(1)
                                    txt_name = Trim(rs_emp_1(2)) 'added by varinder to print dependent name on barcode 09082023
                'SHWETA 30-01-2012
                '                Dim BP1, BP2, BP3 As String
                                If CHK_CASENO = msk_caseno Then
                                Else
                                MsgBox "CHECK CASENO"
                                GoTo BB
                                End If
                                
                
                                BP1 = Left(SCODE111t, 2)
                                BP3 = Right(SCODE111t, 2)
                                BP2 = Mid(SCODE111t, 3, 8)
                                
                
                'FOR TESTING 19-03-2012
                'commented on 25 mar-2012 intermec
                'COMMENTED BY SHWETA ON 28-04-2014
'                PRNSTR = ""
'                                PRNTSTR1 = ""
'                                PRNSTR = "{D0200,0480,0180|}" & vbCrLf
'                                PRNTSTR1 = PRNSTR & "{AY;+10,0|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{C|}" & vbCrLf
'                                'PRNTSTR1 = PRNTSTR1 & "{XB00;0400,0130,A,3,03,2,0050,+0000000000,000,0,00=>6" & Trim(BP1) & ">5" & Trim(BP2) & ">" & Trim(BP3) & "|}" & vbCrLf
'                                '{XB00;0450,0130,A,3,03,2,0050,+0000000000,000,0,00=>6
'                                PRNTSTR1 = PRNTSTR1 & "{XB00;0420,0140,A,3,03,2,0065,+0000000000,000,0,00=>6" & Trim(SCODE111t) & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{PV00;0359,0049,0021,0021,B,22,B=" & Trim(SCODE111t) & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{PV01;0460,0147,0024,0024,B,22,B=" & Trim(txt_name) & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{PV02;0450,0019,0023,0032,B,22,B=" & agesex & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{PV03;0149,0019,0021,0021,B,22,B=" & DESCRIP & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{PV04;0130,0049,0021,0021,B,22,B=" & Trim(txt_status) & "/" & Trim(PCATGS) & "|}" & vbCrLf
'                                'PRNTSTR1 = PRNTSTR1 & "{PV05;0369,0019,0021,0021,B,22,B=" & Trim(msk_caseno) & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{PV05;0355,0016,0025,0032,B,22,B=" & Trim(tblcaseno) & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C2211|}"
                '''''''''''''VARINDER
                
                                PRNTSTR1 = ""
                                PRNSTR = "{D0200,0480,0180|}" & vbCrLf
                                PRNTSTR1 = PRNSTR & "{AY;+10,0|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{C|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{XB00;0375,0140,A,3,02,2,0050,+0000000000,000,0,00=>6" & Trim(SCODE111t) & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{PV00;0359,0070,0021,0021,B,22,B=" & Trim(SCODE111t) & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{PV01;0460,0145,0024,0024,B,22,B=" & Trim(txt_name) & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{PV02;0450,0020,0023,0030,B,22,B=" & agesex & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{PV03;0360,0020,0020,0020,B,22,B=" & DESCRIP & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{PV04;0450,0049,0021,0021,B,22,B=" & Trim(txt_status) & "/" & Trim(PCATGS) & "|}" & vbCrLf
                                'PRNTSTR1 = PRNTSTR1 & "{PV05;0369,0019,0021,0021,B,22,B=" & Trim(msk_caseno) & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{PV05;0355,0040,0025,0030,B,22,B=" & Trim(tblcaseno) & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C2211|}"
                

                
                
                                    Print #f1, PRNTSTR1
                                    
                                    delay = delay + 3
                '                    SELLABNO_RS.MoveNext
                '                Wend
                                colorcd = ""
                                Close #f1
                                Shell bpath & "\micro.bat", 0
                       
BB:
                                SELLABNO_QR.Close
                                
                            Else 'Patient
                                
                                f1 = FreeFile()
                                
                                'Open "c:\disreportsec\data1\micro.TXT" For Output As #f1 'Commneted by Bhushan
                                Open "C:\DISREPORT\BARCODE\micro.txt" For Output As #f1
                                Dim rs_ta101_1 As rdoResultset
                        
                                'STR = "SELECT TESTCODE FROM INTERFACE.mINTERFACE WHERE REQNO=?"
                                'Set SELPROCODE_QR = rdb.CreateQuery("PROCODE", STR)
                                
                                str = "SELECT distinct SAMPLECODE.SAMPLECD FROM DISLIB.SAMPLECODE,DISLIB.DMULPROCE WHERE DMULPROCE.PRO_CODE=SAMPLECODE.SERCODE AND DMULPROCE.REQNO=?"
                                Set SELLABNO_QR = rdb.CreateQuery("SELLABNO", str)
                        
                                SELLABNO_QR(0) = REQNO
                                SELLABNO_QR.Execute
                        
                                c = 0
                                Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
                                
                                c = SELLABNO_RS.RowCount
                                colorcd = ""
                                While Not SELLABNO_RS.EOF
                                
                                    REQ111t = REQNO
                
                                    SCODE111t = SELLABNO_RS(0)
                                    smplcd = SELLABNO_RS(0)
                                    
                                    REQ22 = Mid$(REQ111t, 4)
                                    SCODE111t = REQ22 & SCODE111t
                
                                    deptdesc
                                    
                                    SELLABNO_RS.MoveNext
                                Wend
                
                                    DPT = Mid(REQNO, 1, 3)
                                    
                                    SELDEPTDES_QR(0) = DPT
                                    Set SELDEPTDES_RS = SELDEPTDES_QR.OpenResultset(1, 3)
                                    
                                    DESCRIP = ""
                                    DESCRIP = Mid(Trim(SELDEPTDES_RS("DEPT_NAME")), 1, 4) & "/" & colorcd
                                    
                                    
                                    Set rs_ta101_1 = rdb.OpenResultset("select birth_date,sex,PAT_CATG from patoolslib.ta101 where case_no='" & Trim(msk_caseno.FormattedText) & "'", 1, 1)
                
                                    Dim n_1 As Integer
                                    n_1 = DateDiff("m", rs_ta101_1("BIRTH_DATE"), CStr(DATE11)) / 12
                                    agesex = n_1 & "/" & rs_ta101_1(1)
                                    
                                    'SHWETA 30-01-2012
                '                Dim BP1, BP2, BP3 As String
                                BP1 = Left(SCODE111t, 2)
                                BP3 = Right(SCODE111t, 2)
                                BP2 = Mid(SCODE111t, 3, 8)
                                
                                 If CHK_CASENO = msk_caseno.FormattedText Then
                                Else
                                MsgBox "ELSE CHECK CASENO"
                                GoTo BB
                                End If
                                '''added BY CHARULATA for taking the caseno from rrequisitn on 05-03-2012
                                Set rs_reqnno = rdb.OpenResultset("select caseno from rislib.rrequisitn where reqnno='" & Trim(REQNO) & "'", 1, 1)
                                If rs_reqnno.RowCount > 1 Or rs_reqnno.RowCount = 0 Then
                                
                                    MsgBox ("ERROR in REQUISITION")
                                    
                                ElseIf rs_reqnno.RowCount = 1 Then
                                
                                    tblcaseno = rs_reqnno(0)
                                End If
                       
                           
                       
                       
                'commented on 25 mar-2012 intermec
'                 'COMMENTED ON 16-03-2012 SHWETA 'COMMENTED ON 03-04-2012
'                                   PRNSTR = ""
'                                PRNTSTR1 = ""
'                                PRNSTR = "{D0200,0480,0180|}" & vbCrLf
'                                PRNTSTR1 = PRNSTR & "{AY;+10,0|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{C|}" & vbCrLf
'                                'PRNTSTR1 = PRNTSTR1 & "{XB00;0400,0130,A,3,03,2,0050,+0000000000,000,0,00=>6" & Trim(BP1) & ">5" & Trim(BP2) & ">" & Trim(BP3) & "|}" & vbCrLf
'                                '{XB00;0450,0130,A,3,03,2,0050,+0000000000,000,0,00=>6
'                                PRNTSTR1 = PRNTSTR1 & "{XB00;0375,0130,A,3,02,2,0050,+0000000000,000,0,00=>6" & Trim(SCODE111t) & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{PV00;0359,0049,0021,0021,B,22,B=" & Trim(SCODE111t) & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{PV01;0460,0147,0024,0024,B,22,B=" & Trim(txt_name) & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{PV02;0450,0019,0023,0032,B,22,B=" & agesex & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{PV03;0169,0019,0021,0021,B,22,B=" & DESCRIP & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{PV04;0130,0049,0021,0021,B,22,B=" & Trim(txt_status) & "/" & Trim(PCATGS) & "|}" & vbCrLf
'                                'PRNTSTR1 = PRNTSTR1 & "{PV05;0369,0019,0021,0021,B,22,B=" & Trim(msk_caseno) & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{PV05;0355,0016,0025,0032,B,22,B=" & Trim(tblcaseno) & "|}" & vbCrLf
'                                PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C2211|}"
                                'XB00;0375,0130,A,3,02,2,0050,+0000000000,000,0,00=>6
                                
                                
                                   PRNSTR = ""
                                PRNTSTR1 = ""
                                PRNSTR = "{D0200,0480,0180|}" & vbCrLf
                                PRNTSTR1 = PRNSTR & "{AY;+10,0|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{C|}" & vbCrLf
                                'PRNTSTR1 = PRNTSTR1 & "{XB00;0400,0130,A,3,03,2,0050,+0000000000,000,0,00=>6" & Trim(BP1) & ">5" & Trim(BP2) & ">" & Trim(BP3) & "|}" & vbCrLf
                                '{XB00;0450,0130,A,3,03,2,0050,+0000000000,000,0,00=>6
                                PRNTSTR1 = PRNTSTR1 & "{XB00;0375,0140,A,3,02,2,0050,+0000000000,000,0,00=>6" & Trim(SCODE111t) & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{PV00;0359,0070,0021,0021,B,22,B=" & Trim(SCODE111t) & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{PV01;0460,0145,0024,0024,B,22,B=" & Trim(txt_name) & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{PV02;0450,0020,0023,0030,B,22,B=" & agesex & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{PV03;0360,0020,0020,0020,B,22,B=" & DESCRIP & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{PV04;0450,0049,0021,0021,B,22,B=" & Trim(txt_status) & "/" & Trim(PCATGS) & "|}" & vbCrLf
                                'PRNTSTR1 = PRNTSTR1 & "{PV05;0369,0019,0021,0021,B,22,B=" & Trim(msk_caseno) & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{PV05;0355,0040,0025,0030,B,22,B=" & Trim(tblcaseno) & "|}" & vbCrLf
                                PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C2211|}"
                                
                                    Print #f1, PRNTSTR1
                                    
                                    delay = delay + 3
                '                    SELLABNO_RS.MoveNext
                '                Wend
                                colorcd = ""
                                Close #f1
                                Shell bpath & "\micro.bat", 0
                                
                                SELLABNO_QR.Close
                                
                            End If
                            

            Else 'If other than HL

                 If Len(msk_caseno) = 6 Then 'Staff
                             f1 = FreeFile()
                             Open bpath & "\micro.TXT" For Output As #f1
                             Dim rs_emp As rdoResultset
                             Dim DPT1 As String
                              str = "SELECT distinct SAMPLECODE.SAMPLECD FROM DISLIB.SAMPLECODE,DISLIB.DMULPROCE WHERE DMULPROCE.PRO_CODE=SAMPLECODE.SERCODE AND DMULPROCE.REQNO=?"
                                Set SELLABNO_QR = rdb.CreateQuery("", str)
                
                                SELLABNO_QR(0) = REQNO
                                SELLABNO_QR.Execute
                                c = 0
                                Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
                                c = SELLABNO_RS.RowCount
                                While Not SELLABNO_RS.EOF
                                   'Dim REQ111t, P_CODE111t, SCODE111t
                                    REQ111t = REQNO
                                    SCODE111t = SELLABNO_RS(0)
                                    smplcd = SELLABNO_RS(0)
                                    REQ22 = Mid$(REQ111t, 4)
                                    SCODE111t = REQ22 & SCODE111t
                                     deptdesc
                                                 
                                                 DPT1 = Mid(REQNO, 1, 3)
                                                 
                                                 SELDEPTDES_QR(0) = DPT1
                                                 Set SELDEPTDES_RS = SELDEPTDES_QR.OpenResultset(1, 3)
                                                                         
                                                 DESCRIP = ""
                                                 DESCRIP = Mid(Trim(SELDEPTDES_RS("DEPT_NAME")), 1, 8) & "/" & colorcd
                                                 
                                                 Set rs_emp = rdb.OpenResultset("select age,sex,PNAME from dislib.dreqstaffd where empcode='" & Trim(msk_caseno.FormattedText) & "'", 1, 1)
                                                 agesex1 = rs_emp(0) & "/" & rs_emp(1)
                                                 txt_name = Trim(rs_emp(2))   'added by varinder to print name of the dependent on barcode print
                             'txt_name = Trim(rs_ta101_1("PNAME")) ' & " " & Trim(rs_ta101_1("fIRST_NAME")) & " " & Trim(rs_ta101_1("MID_NAME")) & " " & Trim(rs_ta101_1("lAST_NAME"))
                                                'SHWETA 30-01-2012
                             '                Dim BP1, BP2, BP3 As String
                                             BP1 = Left(SCODE111t, 2)
                                             BP3 = Right(SCODE111t, 2)
                                             BP2 = Mid(SCODE111t, 3, 8)
                                            If Trim(CHK_CASENO) = Trim(msk_caseno) Then
                                             Else
                                             Close #f1
                                             MsgBox "3.CHECK CASENO"
                                             GoTo BB
                                             End If
                                             
                             
                             
'                                          'COMMENTED ON 16-03-2012 SHWETA inter mec
'                                                 PRNSTR = ""
'                                             PRNTSTR1 = ""
'                                             PRNTSTR1 = "{D0200,0480,0180|}" & vbCrLf
'                                             PRNTSTR1 = PRNTSTR1 & "{AY;+10,0|}" & vbCrLf
'                                             PRNTSTR1 = PRNTSTR1 & "{C|}" & vbCrLf
'                                             PRNTSTR1 = PRNTSTR1 & "{XB00;0375,0130,A,3,02,2,0050,+0000000000,000,0,00=>6" & Trim(SCODE111t) & "|}" & vbCrLf
'                                             PRNTSTR1 = PRNTSTR1 & "{PV00;0359,0049,0021,0021,B,22,B=" & Trim(SCODE111t) & "|}" & vbCrLf
'                                             PRNTSTR1 = PRNTSTR1 & "{PV01;0460,0147,0024,0024,B,22,B=" & Trim(txt_name) & "|}" & vbCrLf
'                                             PRNTSTR1 = PRNTSTR1 & "{PV02;0450,0019,0023,0032,B,22,B=" & agesex & "|}" & vbCrLf
'                                             PRNTSTR1 = PRNTSTR1 & "{PV03;0169,0019,0021,0021,B,22,B=" & DESCRIP & "|}" & vbCrLf
'                                             PRNTSTR1 = PRNTSTR1 & "{PV04;0130,0049,0021,0021,B,22,B=" & Trim(txt_status) & "/" & Trim(PCATGS) & "|}" & vbCrLf
'                                             PRNTSTR1 = PRNTSTR1 & "{PV05;0355,0016,0025,0032,B,22,B=" & Trim(msk_caseno.FormattedText) & "|}" & vbCrLf
'                                             PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C2211|}"
                             'XB00;0375,0130,A,3,02,2,0050,+0000000000,000,0,00=>6
                             
                             
                             
                                                 PRNSTR = ""
                                             PRNTSTR1 = ""
                                             PRNTSTR1 = "{D0200,0480,0180|}" & vbCrLf
                                             PRNTSTR1 = PRNTSTR1 & "{AY;+10,0|}" & vbCrLf
                                             PRNTSTR1 = PRNTSTR1 & "{C|}" & vbCrLf
                                             PRNTSTR1 = PRNTSTR1 & "{XB00;0375,0140,A,3,02,2,0050,+0000000000,000,0,00=>6" & Trim(SCODE111t) & "|}" & vbCrLf
                                             PRNTSTR1 = PRNTSTR1 & "{PV00;0359,0070,0021,0021,B,22,B=" & Trim(SCODE111t) & "|}" & vbCrLf
                                             PRNTSTR1 = PRNTSTR1 & "{PV01;0460,0145,0024,0024,B,22,B=" & Trim(txt_name) & "|}" & vbCrLf
                                             PRNTSTR1 = PRNTSTR1 & "{PV02;0450,0020,0023,0030,B,22,B=" & agesex & "|}" & vbCrLf
                                             PRNTSTR1 = PRNTSTR1 & "{PV03;0360,0020,0020,0020,B,22,B=" & DESCRIP & "|}" & vbCrLf
                                             PRNTSTR1 = PRNTSTR1 & "{PV04;0450,0049,0021,0021,B,22,B=" & Trim(txt_status) & "/" & Trim(PCATGS) & "|}" & vbCrLf
                                             PRNTSTR1 = PRNTSTR1 & "{PV05;0355,0040,0025,0030,B,22,B=" & Trim(msk_caseno.FormattedText) & "|}" & vbCrLf
                                             PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C2211|}"
                             
                                                 Print #f1, PRNTSTR1
                                                 'delay = delay + 1 'Shraddha 06/07/09
                                                 delay = delay + 3
                                                 SELLABNO_RS.MoveNext
                                             Wend
                                             
                                             Close #f1
                                             Shell bpath & "\micro.bat", 0
                                    
                                             'SELPROCODE_QR.Close
                                             SELLABNO_QR.Close
                                    'Exit Sub
                                Else 'Patient
                                    'Kill "C:\MICRO.TXT"
                                    f1 = FreeFile()
                                On Error GoTo aopnfile
                                    Open bpath & "\micro.TXT" For Output As #f1
aopnfile:
                                              If Err.Number = 55 Then
                                                Close #f1 - 1
                                                Open "c:\disreportsec\data1\micro.TXT" For Output As #f1
                                              End If
                                    Dim rs_ta101 As rdoResultset
                            
                                    'STR = "SELECT TESTCODE FROM INTERFACE.mINTERFACE WHERE REQNO=?"
                                    'Set SELPROCODE_QR = rdb.CreateQuery("PROCODE", STR)
                                    
                                    str = "SELECT distinct SAMPLECODE.SAMPLECD FROM DISLIB.SAMPLECODE,DISLIB.DMULPROCE WHERE DMULPROCE.PRO_CODE=SAMPLECODE.SERCODE AND DMULPROCE.REQNO=?"
                                    Set SELLABNO_QR = rdb.CreateQuery("", str)
                            
                                    SELLABNO_QR(0) = REQNO
                                    SELLABNO_QR.Execute
                            
                                    c = 0
                                    Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
                                    
                                    c = SELLABNO_RS.RowCount
                    
                                    While Not SELLABNO_RS.EOF
                                        'Dim REQ111t, P_CODE111t, SCODE111t
                                        REQ111t = REQNO
                    
                                        SCODE111t = SELLABNO_RS(0)
                                        smplcd = SELLABNO_RS(0)
                                        
                                        REQ22 = Mid$(REQ111t, 4)
                                        SCODE111t = REQ22 & SCODE111t
                    
                                        deptdesc
                    
                    
                                        DPT = Mid(REQNO, 1, 3)
                                        
                                        SELDEPTDES_QR(0) = DPT
                                        Set SELDEPTDES_RS = SELDEPTDES_QR.OpenResultset(1, 3)
                                        
                                        DESCRIP = ""
                                        DESCRIP = Mid(Trim(SELDEPTDES_RS("DEPT_NAME")), 1, 4) & "/" & colorcd
                                        
                                        
                                        Set rs_ta101 = rdb.OpenResultset("select birth_date,sex, pat_catg from patoolslib.ta101 where case_no='" & Trim(msk_caseno.FormattedText) & "'", 1, 1)
                    
                                        Dim n As Integer
                                        n = DateDiff("m", rs_ta101("BIRTH_DATE"), CStr(DATE11)) / 12
                                        agesex = n & "/" & rs_ta101(1)
                    '                Dim BP1, BP2, BP3 As String
                                    If Trim(CHK_CASENO) = msk_caseno.FormattedText Then
                                    Else
                                    Close #f1
                                    MsgBox "4.CHECK CASENO"
                                    GoTo BB
                                    End If
                                    BP1 = Left(SCODE111t, 2)
                                    BP3 = Right(SCODE111t, 2)
                                    BP2 = Mid(SCODE111t, 3, 8)
                                    '''added BY CHARULATA for taking the caseno from rrequisitn on 05-03-2012
                                    Set rs_reqnno = rdb.OpenResultset("select caseno from rislib.rrequisitn where reqnno='" & Trim(REQNO) & "'", 1, 1)
                                    If rs_reqnno.RowCount > 1 Or rs_reqnno.RowCount = 0 Then
                                    
                                        MsgBox ("ERROR in REQUISITION")
                                        
                                    ElseIf rs_reqnno.RowCount = 1 Then
                                    
                                        tblcaseno = rs_reqnno(0)
                                    End If
                    
                    
                    
                                             
                    
'                                    PRNSTR = ""
'                                    PRNTSTR1 = ""
'                                    PRNTSTR1 = "{D0200,0480,0180|}" & vbCrLf
'                                    PRNTSTR1 = PRNTSTR1 & "{AY;+10,0|}" & vbCrLf
'                                    PRNTSTR1 = PRNTSTR1 & "{C|}" & vbCrLf
'                                    PRNTSTR1 = PRNTSTR1 & "{XB00;0375,0130,A,3,02,2,0050,+0000000000,000,0,00=>6" & Trim(SCODE111t) & "|}" & vbCrLf
'                                    PRNTSTR1 = PRNTSTR1 & "{PV00;0359,0049,0021,0021,B,22,B=" & Trim(SCODE111t) & "|}" & vbCrLf
'                                    PRNTSTR1 = PRNTSTR1 & "{PV01;0460,0147,0024,0024,B,22,B=" & Trim(txt_name) & "|}" & vbCrLf
'                                    PRNTSTR1 = PRNTSTR1 & "{PV02;0450,0019,0023,0032,B,22,B=" & agesex & "|}" & vbCrLf
'                                    PRNTSTR1 = PRNTSTR1 & "{PV03;0169,0019,0021,0021,B,22,B=" & DESCRIP & "|}" & vbCrLf
'                                    PRNTSTR1 = PRNTSTR1 & "{PV04;0130,0049,0021,0021,B,22,B=" & Trim(txt_status) & "/" & Trim(PCATGS) & "|}" & vbCrLf
'                                    PRNTSTR1 = PRNTSTR1 & "{PV05;0355,0016,0025,0032,B,22,B=" & Trim(msk_caseno.FormattedText) & "|}" & vbCrLf
'                                    PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C2211|}"
                                    'XB00;0375,0130,A,3,02,2,0050,+0000000000,000,0,00=>6
                                    
                                    
                                        PRNSTR = ""
                                    PRNTSTR1 = ""
                                    PRNTSTR1 = "{D0200,0480,0180|}" & vbCrLf
                                    PRNTSTR1 = PRNTSTR1 & "{AY;+10,0|}" & vbCrLf
                                    PRNTSTR1 = PRNTSTR1 & "{C|}" & vbCrLf
                                    PRNTSTR1 = PRNTSTR1 & "{XB00;0375,0140,A,3,02,2,0050,+0000000000,000,0,00=>6" & Trim(SCODE111t) & "|}" & vbCrLf
                                    PRNTSTR1 = PRNTSTR1 & "{PV00;0359,0070,0021,0021,B,22,B=" & Trim(SCODE111t) & "|}" & vbCrLf
                                    PRNTSTR1 = PRNTSTR1 & "{PV01;0460,0145,0024,0024,B,22,B=" & Trim(txt_name) & "|}" & vbCrLf
                                    PRNTSTR1 = PRNTSTR1 & "{PV02;0450,0020,0023,0030,B,22,B=" & agesex & "|}" & vbCrLf
                                    PRNTSTR1 = PRNTSTR1 & "{PV03;0360,0020,0020,0020,B,22,B=" & DESCRIP & "|}" & vbCrLf
                                    PRNTSTR1 = PRNTSTR1 & "{PV04;0450,0049,0021,0021,B,22,B=" & Trim(txt_status) & "/" & Trim(PCATGS) & "|}" & vbCrLf
                                    PRNTSTR1 = PRNTSTR1 & "{PV05;0355,0040,0025,0030,B,22,B=" & Trim(msk_caseno.FormattedText) & "|}" & vbCrLf
                                    PRNTSTR1 = PRNTSTR1 & "{XS;I,0001,0002C2211|}"
                                    
                                    
            '        MsgBox PRNTSTR1
                                        Print #f1, PRNTSTR1
                                        'delay = delay + 1 'Shraddha 06/07/09
                                        delay = delay + 3
                                        SELLABNO_RS.MoveNext
                                    Wend
                                    
                                    
                                    
                                    Close #f1
                                    Shell bpath & "\micro.bat", 0
                                    'SELPROCODE_QR.Close
                                    SELLABNO_QR.Close
                                    'Exit Sub
            End If
            
            
            
            End If '17/10/2011
            
            
            MYDELAY (delay)
                
            If newbarcode = True Then
                BARCODECNT
            End If
        End If 'Shraddha 28/05/2009
    End If
    'End If
BB1:
'On Error GoTo ccdd
'Close #f1
'ccdd:
Next i

'If Dir(bpath & "\micro.TXT") <> "" Then
'    Kill bpath & "\micro.TXT" 'Uncommented by Shraddha on 05/05/2009
'End If


'
'    'Dim QTY1 As Integer
'Nos = 0
'J = 1
'f1 = FreeFile()
'
'Open "c:\TATATM.TXT" For Output As #f1
'
'     'FOR NEW FORMAT OF BARCODE PRINTING PLEASE USE
'     'FOLLOWING OPEN COMMAND (JUST REMOVE THE COMMENT) BY MILIND
'     'Open "c:\tata11.TXT" For Output As #f1
'If BARCODE_NEW = True Then
'    SELTIME_QR(0) = reqno
'End If
'If BARCODE_VIEW = True Then
'    SELTIME_QR(0) = II
'End If
'SELTIME_QR.Execute
'Set SELTIME_RS = SELTIME_QR.OpenResultset(1, 3)
'
'If SELTIME_RS.RowCount > 0 Then
'    REC_DATE = Format(SELTIME_RS!SAMRECDATE, "DD/MM/YYYY")
'    REC_TIME = Format(SELTIME_RS!SAMRECTIME, "hh:mm:ss")
'Else
'    MsgBox "YOU CANNOT PRINT BARCODE LABEL FOR THIS REQUISITN B'COZE NO REQNO EXIST IN RTIMELOG FILE"
'    Close #f1
'    Exit Sub
'End If
'
'If BARCODE_NEW = True Then
'
'        STR = "SELECT TESTCODE FROM INTERFACE.DINTERFACE WHERE REQNO=?"
'        Set SELPROCODE_QR = rdb.CreateQuery("PROCODE", STR)
'        STR = "SELECT DISTINCT LABNO FROM INTERFACE.DINTERFACE WHERE REQNO=?"
'        Set SELLABNO_QR = rdb.CreateQuery("LABNO1", STR)
'
'
'        SELLABNO_QR(0) = reqno
'        SELLABNO_QR.Execute
'        QTY = 1
'        Nos = 0
'        Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
'        Do Until SELLABNO_RS.EOF
'            BAR_CD = Trim(SELLABNO_RS!LABNO)
'            SELPATH_QR(0) = reqno
'            SELPATH_QR.Execute
'            Set SELPATH_RS = SELPATH_QR.OpenResultset(1, 3)
'            LABNO = Mid(SELPATH_RS("PATHNO"), 7)
'            Dim count As Integer
'            SELPROCODE_QR(0) = reqno
'            SELPROCODE_QR.Execute
'            Set SELPROCODE_RS = SELPROCODE_QR.OpenResultset(1, 3)
'            count = SELPROCODE_RS.RowCount
'            CASENO = msk_caseno.FormattedText
'            PATNAME = Trim(TXT_NAME.Text)
'            COLL_CENTER = ""
'            QTY1 = "000" & QTY
'            EXP_DT = "EXP" '& Trim(SELLABNO_RS!TDATE) & Trim(SELLABNO_RS!TTIME)
'           'Call Printlbl(BAR_CD)
'            'BY MILIND FOR NEW FORMAT OF BARCODE PRINTING PLEASE USE
'            'FOLLOWING CALL COMMAND (JUST REMOVE THE COMMENT, ALL THE CODING IS DONE)
'            'Call PRINTLBLNEW(BAR_CD)
'            Nos = Nos + QTY
'            SELLABNO_RS.MoveNext
'        Loop
'        Close #f1
'        Shell "c:\TATATM.BAT", 0
'        'FOR NEW FORMAT OF BARCODE PRINTING PLEASE USE
'        'FOLLOWING SHELL COMMAND (REMOVE THE COMMENT) BY MILIND
'        'Shell "c:\tata11.bat", 0
'        MsgBox "Total " & Nos & " Labels Have Been Printed", vbInformation
'
'ElseIf BARCODE_VIEW = True Then
'    If II = "" Then
'    MsgBox "PLEASE CLICK ON REQUISITN NO."
'    Close #f1
'    Exit Sub
'Else
'    SELLABNO_QR(0) = II ' req_chglen(cmb_reqno.Text)
'    SELLABNO_QR.Execute
'    QTY = 1
'    Nos = 0
'    Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
'    Do Until SELLABNO_RS.EOF
'        BAR_CD = Trim(SELLABNO_RS!LABNO)
'        SELPATH_QR(0) = II
'        SELPATH_QR.Execute
'        Set SELPATH_RS = SELPATH_QR.OpenResultset(1, 3)
'        LABNO = Mid(SELPATH_RS("PATHNO"), 7)
'        CASENO = msk_caseno.FormattedText
'        PATNAME = Trim(TXT_NAME.Text)
'        COLL_CENTER = ""
'        QTY1 = "000" & QTY
'        EXP_DT = "EXP" '& Trim(SELLABNO_RS!TDATE) & Trim(SELLABNO_RS!TTIME)
'        '(BAR_CD)
'        Nos = Nos + QTY
'        SELLABNO_RS.MoveNext
'    Loop
'End If
'    Close #f1
'    Shell "c:\TATATM.bat", 0
'    MsgBox "Total " & Nos & " Labels Have Been Printed", vbInformation
'    End If
'    Close #f1
'
End Sub
Public Sub deptdesc()
    If Mid(REQNO, 4, 2) = "HL" Then
    
    Else
    colorcd = ""
    End If

    If smplcd = "A" Or smplcd = "B" Then
        colorcd = colorcd & " RED"
    ElseIf smplcd = "R" Or smplcd = "S" Or smplcd = "P" Then
        colorcd = colorcd & " GRAY"
    ElseIf smplcd = "E" Then
        colorcd = colorcd & " E.P"
    ElseIf smplcd = "O" Then
        colorcd = colorcd & " VITAMIN"
    ElseIf smplcd = "M" Then
        colorcd = colorcd & " MARKER"
    ElseIf smplcd = "T" Then
        colorcd = colorcd & " THYRO"
    ElseIf smplcd = "W" Then
        colorcd = colorcd & " 24URINE"
    ElseIf smplcd = "U" Then
        colorcd = colorcd & " Coag"
    ElseIf smplcd = "D" Then
        colorcd = colorcd & " ESR"
    ElseIf smplcd = "I" Then
        colorcd = colorcd & " Retic"
    ElseIf smplcd = "C" Then
        colorcd = colorcd & " CBC"
    ElseIf smplcd = "Q" Then
        colorcd = colorcd & " MDC"
    ElseIf smplcd = "W" Then
        colorcd = colorcd & " MPC"
    ElseIf smplcd = "Y" Then
        colorcd = colorcd & " PBS"
    ElseIf smplcd = "B" Then
        colorcd = colorcd & " METHOTR"
    ElseIf smplcd = "Z" Then
        colorcd = colorcd & " URINE"
    ElseIf smplcd = "V" Then
        colorcd = colorcd & " CSF"
    ElseIf smplcd = "Y" Then
        colorcd = colorcd & " FLUID"
    Else
        colorcd = ""
    End If
End Sub

Private Sub Form_Load()
On Error GoTo MSG:
NEWFLAG = False
viewflag = False
newbarcode = False
allflag = False
casenoflag = False
datetime
txt_date.Text = DATE11
txt_time.Text = time1
chk_collected.Visible = False
btnenabl ("INI")
rlocationflag = False

If COLLECTIONFLAG = False Then
''*****ADDED BY MONO ON 30/10/03
str = "SELECT PATHNO,REFTODOC FROM DISLIB.COLLECTLF WHERE REQNNO=?"
Set SELPATHNO_QR = rdb.CreateQuery("SELPATHNO_STR", str)


''*****ADDED BY MONO ON 30/10/03
   ''SELECT DATA FROM BLOODPRODU
    str = "SELECT * FROM DISLIB.BLOODPRODU WHERE REQNO=?"
    Set SELBLOODPRODU_QR = rdb.CreateQuery("SELBLOODPRODU_STR11", str)

str = "select * from RISLIB.rlocation where dept_code=? and Modalitycd=?"
Set rlocation_qr = rdb.CreateQuery("RLOCATION11_QR1", str)

'for selecting all collection centres
'ADD MONO ON 06/08/03
str = "select DISTINCT(LOCATION) from RISLIB.rlocation WHERE (DEPT_CODE <> 'LZZ' OR DEPT_CODE <> 'WAA') ORDER BY LOCATION"
Set rlocation_qrall = rdb.CreateQuery("RLOCATION11_QR1all", str)

str = "SELECT * FROM PATOOLSLIB.MA106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR  DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ') "
Set SELMA106_QR = rdb.CreateQuery("SELMA1061_QRcc", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE (REQNDT BETWEEN ? AND ?) AND (REQNSTAT= ? OR REQNSTAT=?) AND (DEPT_CODE=? AND MODALITYCD=?) AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'CAZ' AND DEPT_CODE <> 'WAA')"
'STR = "SELECT * FROM DISLIB.COLLECTVIEW WHERE REQNDT BETWEEN ? AND ?  "
Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN_QR12", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND (REQNSTAT= ? OR REQNSTAT=?) AND DEPT_CODE=? AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'CAZ' AND DEPT_CODE <> 'WAA') order by caseno"
'STR = "SELECT * FROM DISLIB.COLLECTVIEW WHERE REQNDT BETWEEN ? AND ?  "
Set SELRREQUISITNALL_QR = rdb.CreateQuery("SELRREQUISITNALL_QR12", str)


'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND (REQNSTAT= ? OR REQNSTAT=?) AND DEPT_CODE=? AND MODALITYCD=? AND LOCATION = ?"
'STR = "SELECT * FROM DISLIB.COLLECTVIEW WHERE REQNDT BETWEEN ? AND ?  "
str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND DEPT_CODE=? AND MODALITYCD=? AND LOCATION = ?"
Set SELRREQUISITN_QRVIEW = rdb.CreateQuery("SELRREQUISITN_QR12VIEW", str)

'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND (REQNSTAT= ? OR REQNSTAT=?) AND DEPT_CODE=? AND LOCATION =?"
'STR = "SELECT * FROM DISLIB.COLLECTVIEW WHERE REQNDT BETWEEN ? AND ?  "
str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ?  AND DEPT_CODE=? AND LOCATION =?"
Set SELALLMODALITY_QRVIEW = rdb.CreateQuery("SELALLMODALITY_QRVIEW", str)

'Added by milind on 22-02-2006 for general opd and private opd
'str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND REQNSTAT=? AND PATNSTAT='OP' AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ') ORDER BY CASENO"
'added by shweta  for investigation appointment 09-04-2014
str = "SELECT * FROM RISLIB.RREQUISITN WHERE Apptdt BETWEEN ? AND ? AND REQNSTAT=? AND PATNSTAT='OP' AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ') ORDER BY CASENO"
Set selrreqn_opqrall = rdb.CreateQuery("selrreqn_opqrall", str)


'------------------------------- Shraddha Pandit 06/04/09 ---------------------------
'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNDT BETWEEN ? AND ?) AND REQNSTAT=? AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ'AND DEPT_CODE <> 'WAB')ORDER BY CASENO"
'str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNDT BETWEEN ? AND ?) AND (REQNSTAT= ? OR REQNSTAT=?) AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ'AND DEPT_CODE <> 'WAB')ORDER BY CASENO"
'*********** Above commented and next added by Shraddha on 16/07/09 ****************
'str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNDT BETWEEN ? AND ?) AND (REQNSTAT= ? OR REQNSTAT=?) AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ'AND DEPT_CODE <> 'WAB' and DEPT_CODE <> 'NZZ')ORDER BY CASENO"

'Commented and added by Shraddha on 19/08/10
'str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNDT BETWEEN ? AND ?) AND (REQNSTAT= ? OR REQNSTAT=?) AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ'AND DEPT_CODE <> 'WAB' and DEPT_CODE <> 'NZZ' and MODALITYCD <> 'FA' AND MODALITYCD <> 'ML' and MODALITYCD <> 'SP' and MODALITYCD <> 'CN')ORDER BY CASENO"
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNDT BETWEEN ? AND ?) AND (REQNSTAT= ? OR REQNSTAT=? ) AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ'AND DEPT_CODE <> 'WAB' and MODALITYCD <> 'FA' AND MODALITYCD <> 'ML' and MODALITYCD <> 'SP' and MODALITYCD <> 'CN')ORDER BY CASENO"
'*********** Above commented and next added by SHWETA on 03-NOV-09 ****************
'str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNDT BETWEEN ? AND ?) AND (REQNSTAT= ? OR REQNSTAT=?) AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ'AND DEPT_CODE <> 'WAB' and DEPT_CODE <> 'NZZ') AND REQNNO IN (SELECT REQNO FROM DISLIB.BARCODE) ORDER BY CASENO"

Set selrreqn_opqr = rdb.CreateQuery("selrreqn_opqr", str)
'-------------------------------------------------------------------------------------

'***************** Shraddha Pandit 16/07/2009 ******************
'str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNDT BETWEEN ? AND ?) AND DEPT_CODE = 'NZZ' and MODALITYCD='SS'"
'Commented and added by Shraddha on 08/02/10
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNDT BETWEEN ? AND ?) AND DEPT_CODE = 'XNZZ' and (MODALITYCD='SS' OR MODALITYCD='BP') AND (REQNSTAT= 'P' OR REQNSTAT='C')"
'str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNDT BETWEEN ? AND ?) AND DEPT_CODE = 'NZZ' and (MODALITYCD='SS' OR MODALITYCD='BP') AND (REQNSTAT= 'P' OR REQNSTAT='C')"
Set sel_ssrreqn_qr = rdb.CreateQuery("", str)
'***************************************************************

'***************** Shraddha Pandit 28/05/2009 ******************
str = "select LABELS from dislib.barcode where reqno=?"
Set sel_barlbl_qr = rdb.CreateQuery("", str)
VF = False '28/05/2009
'***************************************************************



''''''added by shweta as told by kalpana ma'm to display data for all status
'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNDT BETWEEN ? AND ?)  AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ'AND DEPT_CODE <> 'WAB')ORDER BY CASENO"
'Set selrreqn_opqr = rdb.CreateQuery("selrreqn_opqr", STR)

'upto here


str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND REQNSTAT= ? AND (DEPT_CODE <>'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ') AND CASENO IN(SELECT CASE_NO FROM PATOOLSLIB.TA105 WHERE DISC_DATE IS NULL AND WARD_NAME=?) ORDER BY CASENO"
'STR = "SELECT * FROM DISLIB.COLLECTVIEW WHERE REQNDT BETWEEN ? AND ?  "
Set SELRREQUISITN_QRALL = rdb.CreateQuery("SELRREQUISITN_QR12ALL", str)


''SELECT ALL CASES OF ALL DEPARTMENT FOR A PARTICULAR CASENO
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNSTAT= ? OR REQNSTAT=?)AND (DEPT_CODE <>'LZZ' or DEPT_CODE <> 'WAA' )"
Set SELRREQUISITN_QRCASE = rdb.CreateQuery("SELRREQUISITN_QR12ALLCASE", str)


'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND (REQNSTAT= ? OR REQNSTAT=?) AND LOCATION =?  "
'STR = "SELECT * FROM DISLIB.COLLECTVIEW WHERE REQNDT BETWEEN ? AND ?  "
str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND LOCATION =? and (DEPT_CODE <> 'LZZ' and DEPT_CODE <> 'WAA') "
Set SELRREQUISITN_QRALLVIEW = rdb.CreateQuery("SELRREQUISITN_QR12ALLVIEW", str)


str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
Set SELMA101_QR = rdb.CreateQuery("SELMA101_QR11", str)


str = "UPDATE RISLIB.RREQUISITN SET LOCATION=?, REQNSTAT=?,APNTDT=? WHERE REQNNO=? AND (DEPT_CODE <> 'LZZ' or dept_code <> 'WAA') "
Set updrrequisitn_qr = rdb.CreateQuery("UPDREQUISITN_QR1", str)

 str = "UPDATE RISLIB.RTIMELOG SET SAMCOLDATE=?,SAMCOLTIME=? WHERE REQNNO=?"
 Set UPDRTIMECN_QR = rdb.CreateQuery("QWEQWEQWECN", str)

'STR = "INSERT INTO RISLIB.RTIMELOG(SAMCOLDATE,SAMCOLTIME,REQNNO) VALUES(?,?,?)"
'Set UPDRTIMECN_QR = rdb.CreateQuery("QWEQWEQWECN", STR)


' DONE BY ALBERT
str = "SELECT * FROM RISLIB.RREQUISITN WHERE (REQNSTAT= ? OR REQNSTAT=?) AND (REQNDT BETWEEN ? AND ? ) AND CASENO IN(SELECT CASE_NO FROM PATOOLSLIB.TA105 WHERE DISC_DATE IS NULL) ORDER BY CASENO "
 'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNSTAT=? OR REQNSTAT =?) AND DEPT_CODE=? AND MODALITYCD=? "
' UP TO
Set SELREQNO_QR = rdb.CreateQuery("SELREQNO_QR1", str)

'added by milind on 22-02-2006 for staff
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNSTAT= ? OR REQNSTAT=?) AND (REQNDT BETWEEN ? AND ? )AND PATNSTAT='ST' AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ') order by modalitycd,reqndt desc"
Set SELREQNO_STFQR = rdb.CreateQuery("SELREQNO_STFQR", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE (REQNDT BETWEEN ? AND ?) AND (REQNSTAT= ? OR REQNSTAT=?) AND PATNSTAT='ST'  AND (DEPT_CODE <>'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ')ORDER BY CASENO"
Set SELREQNO_STFQRALL = rdb.CreateQuery("SELREQNO_STFQRALL", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND (REQNSTAT= ? OR REQNSTAT=?) AND DEPT_CODE=? AND PATNSTAT='ST' AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'CAZ' AND DEPT_CODE <> 'WAA')ORDER BY CASENO"
Set SELRREQUISITNALL_STFQR = rdb.CreateQuery("SELRREQUISITNALL_STFQR", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE (REQNDT BETWEEN ? AND ?) AND (REQNSTAT= ? OR REQNSTAT=?) AND (DEPT_CODE=? AND MODALITYCD=?) AND PATNSTAT='ST' AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'CAZ' AND DEPT_CODE <> 'WAA')ORDER BY CASENO"
Set SELRREQUISITN_STFQR = rdb.CreateQuery("SELRREQUISITN_STFQR", str)

''''KANCHAN
str = "SELECT SAMRECDATE,SAMRECTIME FROM RISLIB.RTIMELOG WHERE REQNNO=?"
Set SELTIME_QR = rdb.CreateQuery("FFHSHF", str)

str = "SELECT PATHNO FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELPATH_QR = rdb.CreateQuery("PATHNO", str)

str = "SELECT * FROM PATOOLSLIB.MA106 WHERE DEPT_CODE=?"
Set SELDEPTDES_QR = rdb.CreateQuery("DEPTDESC", str)

str = "INSERT INTO DISLIB.BARCODE (REQNO,CASENO,LABELS,PID,DATE,TIME,LOCATION) VALUES (?,?,?,?,?,?,?)"
Set INS_BARCODE = rdb.CreateQuery("INSERTBARCODE", str)
'UPTO HERE

'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNSTAT=? OR REQNSTAT =?) AND (REQNDT BETWEEN ? AND ? )AND (LOCATION =? OR LOCATION=?)"
'Commented and added by Shraddha Pandit on 02/02/09
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNSTAT=? OR REQNSTAT =?) AND (REQNDT BETWEEN ? AND ? )" 'AND (LOCATION =? OR LOCATION=?)"
'Commented and added by SHWETA on 03/10/09
'str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNSTAT=? OR REQNSTAT =?) AND (REQNDT BETWEEN ? AND ? ) AND REQNNO IN (SELECT REQNO FROM DISLIB.BARCODE)" 'AND (LOCATION =? OR LOCATION=?)"
'Commented and added by Shraddha on 01/11/2011
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNSTAT=? OR REQNSTAT =?) AND (REQNDT BETWEEN ? AND ? ) " 'AND REQNNO IN (SELECT REQNO FROM DISLIB.BARCODE)"
Set SELREQNO_QRVIEW = rdb.CreateQuery("SELREQNO_QR1VIEW", str)


str = "select * from patoolslib.ta101 where case_no=?"
Set selta101_qr1 = rdb.CreateQuery("selta101_qr1q", str)

'str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE EMPCODE=?" '--- Commented and added below query to fetch the staff latest reqnstat details
str = "SELECT * FROM RISLIB.RREQUISITN A JOIN DISLIB.DREQSTAFFD B ON A.REQNNO=B.REQNO WHERE A.CASENO=?  AND REQNSTAT='P' "
Set SELDREQSTAFFD_QR = rdb.CreateQuery("SELDREQSTAFFD_QR1", str)

str = "SELECT * FROM RISLIB.RMODALITY WHERE DEPT_CODE=?"
Set SELRMODLAITYQR = rdb.CreateQuery("SELRMODLAITYQRCC", str)

str = "select * from dislib.dmulproce where caseno=? and collection is null"
Set SELDMULPROCEQR = rdb.CreateQuery("seldmulproceqrcc", str)

str = "UPDATE DISLIB.DMULPROCE SET COLLECTION=? WHERE PRO_CODE=? AND REQNO=?"
Set upddmulproceqr = rdb.CreateQuery("upddmulproceqrBC", str)

str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNDT BETWEEN ? AND ? AND COTTECTION IS NULL"
Set SELDMULPROCEQRDT = rdb.CreateQuery("SELDMULPROCEQRDTCC", str)

str = "SELECT HDAY FROM PATOOLSLIB.ma1121 where HDAY= ?"
Set LAST_W_QR = rdb.CreateQuery("LAST_W_QR", str)
LAST_W_FLG = False
P_LAST_WORKING_DAY = ""
lcount = 0
hday = 0
mydate = DATE11
While lcount < 1
   If lcount < 1 Then
    P_LAST_WORKING_DAY = DateAdd("d", -375, mydate)
   End If
   ldate = Format(CDate(P_LAST_WORKING_DAY), "YYYY-MM-DD")
     LAST_W_QR(0) = ldate
     If LAST_W_FLG = False Then
       Set LAST_W_RS = LAST_W_QR.OpenResultset(1, 3)
       LAST_W_FLG = True
     Else
       LAST_W_RS.MoreResults
       LAST_W_RS.Requery
     End If

     If LAST_W_RS.RowCount > 0 Then
      hday = 1
     End If
   If (Weekday(CDate(P_LAST_WORKING_DAY)) <> 1) And hday = 0 Then

    lcount = lcount + 1
   End If
   hday = 0
   mydate = DateAdd("d", -375, mydate)
Wend
If rlocationflagall = False Then
        Set rlocation_rsall = rlocation_qrall.OpenResultset(1, 3)
        rlocationflagall = True
   Else
        rlocation_rsall.MoreResults
        rlocation_rsall.Requery
    End If
    If rlocation_rsall.RowCount > 0 Then
        cmb_location.clear
        While Not rlocation_rsall.EOF
            cmb_location.AddItem Trim(rlocation_rsall("LOCATION"))
            rlocation_rsall.MoveNext
        Wend
    End If


COLLECTIONFLAG = True
End If
MSFlexGrid1.row = 1
MSFlexGrid1.col = 3
chk_collected.Width = MSFlexGrid1.CellWidth
chk_collected.Height = MSFlexGrid1.CellHeight
chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'TXT_FILMQTY.Width = MSF.CellWidth
'TXT_FILMQTY.Height = MSF.CellHeight
'TXT_FILMQTY.Left = MSF.CellLeft + MSF.Left
'TXT_FILMQTY.Top = MSF.CellTop + MSF.Top
cmb_location.Enabled = False
msk_fromdt.Enabled = False
msk_todt.Enabled = True
OPT_ALL.Enabled = False
OPT_CASENO.Enabled = False
Opt_patient.Enabled = False
opt_staff.Enabled = False
msk_caseno.Enabled = False
cmb_modality.Enabled = False
cmb_dept.Enabled = False
cmd_print1.Visible = False
'cmd_scanner.Enabled = False

'***** Shraddha 27/09/2011 *****
'Set location by default according to IP address
If Trim(ipadd) = "10.2.99.103" Then
    cmb_location.Text = "GEN.OPD6"
End If

If Trim(ipadd) = "10.2.1.67" Or Trim(ipadd) = "10.2.1.68" Then
    cmb_location.Text = "GEN.OPD"
End If

If Trim(ipadd) = "10.3.100.57" Then
    cmb_location.Text = "PVT.OPD"
End If
'*******************************

msk_fromdt.Text = P_LAST_WORKING_DAY
msk_fromdt.Enabled = False
msk_todt.Enabled = False
Exit Sub
MSG:
MsgBox Err.Description
End Sub
Public Sub btnenabl(cap As String)

If cap = "NEW" Then
    'cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    'cmd_modify.Enabled = False
'    cmd_print.Enabled = False
    cmd_view.Enabled = False
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    'cmd_modify.Enabled = True
    'cmd_modify.Enabled = False
    CMD_NEW.Enabled = False
    'cmd_save.Enabled = False
   ' cmd_print.Enabled = True
    cmd_view.Enabled = False
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    'cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    'cmd_save.Enabled = False
    'cmd_print.Enabled = False
    cmd_view.Enabled = True
End If


End Sub


Private Sub MSFlexGrid1_Click()
'On Error GoTo MSG:
 
    row1 = MSFlexGrid1.row

    reqvw = req_chglen(MSFlexGrid1.TextMatrix(row1, 2))
 
     MSFlexGrid1.col = 3
    'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then

            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            chk_collected = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))
            
            chk_collected.SetFocus
            
            chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            
            If Trim(MSFlexGrid1.Text) = "Print Barcode" Then
                chk_collected.Value = 1
            Else
                chk_collected.Value = 0
            End If
            
    'End If
    Exit Sub
    
MSG:
    MsgBox Err.Description
End Sub

Private Sub msk_caseno_LostFocus()
'On Error GoTo MSG:
If Len(Trim(msk_caseno.Text)) = 0 Then Exit Sub
If Opt_patient.Value = True Then ''01/12/2022
''' 1/12/2022''''''''''''
'    CASENO = padzero(5, "C", msk_caseno.Text, 3)
'    msk_caseno.Text = UCase(CASENO)
'''''''''''till here(1/12/2022)''''''''''
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
''''''''''''''''''''till here (01/12/2022)'''''''''
    If selta101flag1 = False Then
        selta101_qr1(0) = msk_caseno.FormattedText
        Set selta101_rs1 = selta101_qr1.OpenResultset(1, 3)
        selta101flag1 = True
    Else
        selta101_qr1(0) = msk_caseno.FormattedText
        selta101_rs1.MoreResults
        selta101_rs1.Requery
    End If
    If selta101_rs1.RowCount > 0 Then
        txt_name.Text = Trim(selta101_rs1("tag")) & " " & Trim(selta101_rs1("first_name")) & " " & Trim(selta101_rs1("mid_name")) & " " & Trim(selta101_rs1("last_name"))
    End If
    msk_fromdt.Enabled = True
    msk_fromdt.SetFocus
    getPhoto
ElseIf opt_staff.Value = True Then
    STAFF
    
End If

''***** Added by Shraddha 21/10/2011 *****
'If OPT_CASENO.Value = True And Opt_patient.Value = True Then
'    Call ret_pat_data
'ElseIf OPT_CASENO.Value = True And opt_staff = True Then
'    Call ret_staff_data
'End If
''*****************************************


Exit Sub
MSG:
MsgBox Err.Description
End Sub
Public Sub MapPhoto()
'----------07-03-2014
On Error Resume Next
Dim a As New CNetSrv
Dim o As Boolean
Dim objfso As New FileSystemObject
If objfso.DriveExists("K:") Then
    a.disconnetDrive "K:"
End If
    o = a.mapDrive("K:", "\\10.80.1.9\PATPHOTO", "mapping", "iseries@13")
End Sub

Sub getPhoto()
On Error Resume Next
MapPhoto
Image1.Picture = LoadPicture("")
Image1.Picture = LoadPicture("K:\" & Trim(msk_caseno.Text) & ".jpg")
'Picture1.Picture = LoadPicture("\\10.80.1.9\patphoto\" & Trim(msk_caseno.Text) & ".jpg")
End Sub
Private Sub msk_fromdt_LostFocus()
On Error GoTo MSG:
'********** Shraddha 06/07/09 ***********
If msk_fromdt.FormattedText = mydate Then
Else
    NEWFLAG = True
End If
'****************************************
If IsDate(GET_REVDATE(msk_fromdt.FormattedText)) = False Then
    MsgBox "Pleae ENTER the correct DATE."
    msk_fromdt.Mask = ""
    msk_fromdt.Text = ""
    msk_fromdt.Mask = "##/##/####"
'   MaskEd_from_date.Text = ""
    msk_fromdt.SetFocus
    'Exit Sub
Else
    msk_todt.Enabled = True
    msk_todt.SetFocus
End If
Exit Sub

MSG:
MsgBox Err.Description
End Sub
Private Sub msk_todt_LostFocus()

If OPT_CASENO.Value = True And Opt_patient.Value = True Then
    Call ret_pat_data
ElseIf OPT_CASENO.Value = True And opt_staff = True Then
    Call ret_staff_data
End If
'Commented on 21/10/2011 by Shraddha

'If OPT_CASENO.Value = True And Opt_patient.Value = True Then
'    Call RETRIVE_DATA
'ElseIf OPT_CASENO.Value = True And opt_staff = True Then 'Or (cmb_location.Text = "STAFF") Then
'    Call RETRIV_STAFF_DATE
'End If

'MSG:
'MsgBox Err.Description
End Sub

Public Sub ret_staff_data()
If NEWFLAG = True Then
    If (Trim(cmb_location.Text) = "STAFF") Or (OPT_CASENO.Value = True) Then
        NEWFLAG = False
        SELREQNO_STFQR(0) = msk_caseno.FormattedText
        SELREQNO_STFQR(1) = "P"
        SELREQNO_STFQR(2) = "C"
        SELREQNO_STFQR(3) = GET_REVDATE(msk_fromdt.FormattedText)
        SELREQNO_STFQR(4) = GET_REVDATE(msk_todt.FormattedText)
        'SELREQNO_STFQR(4) = Trim(cmb_location.Text)

        If SELREQNOSTFFLAG = False Then
            Set SELREQNO_STFRS = SELREQNO_STFQR.OpenResultset(1, 3)
            SELREQNOSTFFLAG = True
        Else
            SELREQNO_STFRS.MoreResults
            SELREQNO_STFRS.Requery
        End If

        If SELREQNO_STFRS.RowCount > 0 Then
            i = 1
            MSFlexGrid1.Rows = 1
            While Not SELREQNO_STFRS.EOF
                'If Trim(msk_caseno.FormattedText) = SELREQNO_STFQR(0) Then
                If SELPATHNO_FLAG = False Then
                    SELPATHNO_QR(0) = SELREQNO_STFRS("REQNNO")
                    Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                    SELPATHNO_FLAG = True
                Else
                    SELPATHNO_QR(0) = SELREQNO_STFRS("REQNNO")
                    SELPATHNO_RS.MoreResults
                    SELPATHNO_RS.Requery
                End If
                If Trim(SELREQNO_STFRS("ser_code")) <> "8888" Then
                    If SELMA101FLAG = False Then
                        SELMA101_QR(0) = SELREQNO_STFRS("SER_CODE")
                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                        SELMA101FLAG = True
                    Else
                        SELMA101_QR(0) = SELREQNO_STFRS("SER_CODE")
                        SELMA101_RS.MoreResults
                        SELMA101_RS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        MSFlexGrid1.AddItem i & vbTab & SELREQNO_STFRS("CASENO") & vbTab & req_chglen(SELREQNO_STFRS("REQNNO"))
                        i = i + 1
                    End If
                Else
                    MSFlexGrid1.AddItem i & vbTab & SELREQNO_STFRS("CASENO") & vbTab & req_chglen(SELREQNO_STFRS("REQNNO"))
                    i = i + 1
                End If
                SELREQNO_STFRS.MoveNext
                    'Else
                       ' SELREQNO_STFRS.MoveNext
                    'End If
            Wend
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 3
            chk_collected.Width = MSFlexGrid1.CellWidth
            chk_collected.Height = MSFlexGrid1.CellHeight
            chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_collected.Value = 0
            chk_collected.Visible = True
        End If
    Else
        If Trim(cmb_location.Text) = "STAFF" And Trim(cmb_dept.Text) = "ALL" Then
            NEWFLAG = False
            'SELREQNO_QR(0) = msk_caseno.FormattedText
            SELREQNO_STFQRALL(0) = "P"
            SELREQNO_STFQRALL(1) = "P"
            SELREQNO_STFQRALL(2) = GET_REVDATE(msk_fromdt.FormattedText)
            SELREQNO_STFQRALL(3) = GET_REVDATE(msk_todt.FormattedText)
            'SELREQNO_STFQR(4) = Trim(cmb_location.Text)
            If SELREQNOSTFFLAGALL = False Then
                Set SELREQNO_STFRSALL = SELREQNO_STFQRALL.OpenResultset(1, 3)
                SELREQNOSTFFLAGALL = True
            Else
                SELREQNO_STFRSALL.MoreResults
                SELREQNO_STFRSALL.Requery
            End If

            If SELREQNO_STFRSALL.RowCount > 0 Then
                i = 1
                MSFlexGrid1.Rows = 1
                While Not SELREQNO_STFRSALL.EOF
                    If SELPATHNO_FLAG = False Then
                        SELPATHNO_QR(0) = SELREQNO_STFRSALL("REQNNO")
                        Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                        SELPATHNO_FLAG = True
                    Else
                        SELPATHNO_QR(0) = SELREQNO_STFRSALL("REQNNO")
                        SELPATHNO_RS.MoreResults
                        SELPATHNO_RS.Requery
                    End If
                    If SELMA101FLAG = False Then
                        SELMA101_QR(0) = SELREQNO_STFRSALL("SER_CODE")
                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                        SELMA101FLAG = True
                    Else
                        SELMA101_QR(0) = SELREQNO_STFRSALL("SER_CODE")
                        SELMA101_RS.MoreResults
                        SELMA101_RS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        MSFlexGrid1.AddItem i & vbTab & SELREQNO_STFRSALL("CASENO") & vbTab & req_chglen(SELREQNO_STFRSALL("REQNNO"))
                        i = i + 1
                    End If
                    SELREQNO_STFRSALL.MoveNext
                Wend
                MSFlexGrid1.row = 1
                MSFlexGrid1.col = 3
                chk_collected.Width = MSFlexGrid1.CellWidth
                chk_collected.Height = MSFlexGrid1.CellHeight
                chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                chk_collected.Value = 0
                chk_collected.Visible = True
            End If
        End If
    End If
    
    If (Trim(cmb_dept.Text) <> "ALL") And (cmb_location.Text = "STAFF") Then
        If Trim(cmb_modality.Text) = "" Then
            NEWSTFALLMODALITY
        Else
            If SELRREQUISITN_STFFLAG = False Then
                SELRREQUISITN_STFQR(0) = msk_fromdt.FormattedText
                SELRREQUISITN_STFQR(1) = msk_todt.FormattedText
                SELRREQUISITN_STFQR(2) = "P"
                SELRREQUISITN_STFQR(3) = "P"
                SELRREQUISITN_STFQR(4) = dept
                i = InStr(cmb_modality.Text, "$")
                SELRREQUISITN_STFQR(5) = Mid(cmb_modality, i + 1, 2)
                'SELRREQUISITN_QR(6) = Trim(cmb_location.Text)
                Set SELRREQUISITN_STFRS = SELRREQUISITN_STFQR.OpenResultset(1, 3)
            Else
                SELRREQUISITN_STFQR(0) = msk_fromdt.FormattedText
                SELRREQUISITN_STFQR(1) = msk_todt.FormattedText
                SELRREQUISITN_STFQR(2) = "P"
                SELRREQUISITN_STFQR(3) = "P"
                SELRREQUISITN_STFQR(4) = dept
                i = InStr(cmb_modality.Text, "$")
                SELRREQUISITN_STFQR(5) = Mid(cmb_modality, i + 1, 2)
                'SELRREQUISITN_QR(6) = Trim(cmb_location.Text)
                SELRREQUISITN_STFRS.MoreResults
                SELRREQUISITN_STFRS.Requery
            End If
            NEWFLAG = False
            If SELRREQUISITN_STFRS.RowCount > 0 Then
                txt_rowcount.Text = "Total Number of Requisition for Which Sample have to Collect    " & SELRREQUISITN_STFRS.RowCount
                i = 1
                MSFlexGrid1.Rows = 1
                While Not SELRREQUISITN_STFRS.EOF
                    If SELPATHNO_FLAG = False Then
                        SELPATHNO_QR(0) = SELRREQUISITN_STFRS("REQNNO")
                        Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                        SELPATHNO_FLAG = True
                    Else
                        SELPATHNO_QR(0) = SELRREQUISITN_STFRS("REQNNO")
                        SELPATHNO_RS.MoreResults
                        SELPATHNO_RS.Requery
                    End If
                    If Trim(SELRREQUISITN_STFRS("ser_code")) <> "8888" Then
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = SELRREQUISITN_STFRS("SER_CODE")
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = SELRREQUISITN_STFRS("SER_CODE")
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_STFRS("CASENO") & vbTab & req_chglen(SELRREQUISITN_STFRS("REQNNO"))
                            i = i + 1
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_STFRS("CASENO") & vbTab & req_chglen(SELRREQUISITN_STFRS("REQNNO"))
                        i = i + 1
                    End If
                    SELRREQUISITN_STFRS.MoveNext
                Wend
                MSFlexGrid1.row = 1
                MSFlexGrid1.col = 3
                chk_collected.Width = MSFlexGrid1.CellWidth
                chk_collected.Height = MSFlexGrid1.CellHeight
                chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                chk_collected.Value = 0
                chk_collected.Visible = True
            Else
                MsgBox "There are No Pending Requisition No."
            End If
        End If
    End If
End If
    
If viewflag = True Then
    viewflag = FLASE
    If SELREQNOFLAG = False Then
        SELREQNO_STFQR(0) = Trim(msk_caseno.FormattedText)
        SELREQNO_STFQR(1) = "C"
        SELREQNO_STFQR(2) = "C"
        SELREQNO_STFQR(3) = GET_REVDATE(msk_fromdt.FormattedText)
        SELREQNO_STFQR(4) = GET_REVDATE(msk_todt.FormattedText)
        'SELREQNO_STFQR(4) = Trim(cmb_location.Text)
        Set SELREQNO_STFRS = SELREQNO_STFQR.OpenResultset(1, 3)
        SELREQNOFLAG = True
    Else
        SELREQNO_STFQR(0) = msk_caseno.FormattedText
        SELREQNO_STFQR(1) = "P"
        SELREQNO_STFQR(2) = "C"
        SELREQNO_STFQR(3) = GET_REVDATE(msk_fromdt.FormattedText)
        SELREQNO_STFQR(4) = GET_REVDATE(msk_todt.FormattedText)
        'SELREQNO_STFQR(4) = Trim(cmb_location.Text)
        SELREQNO_STFRS.MoreResults
        SELREQNO_STFRS.Requery
    End If
    If SELREQNO_STFRS.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELREQNO_STFRS.EOF
            If SELPATHNO_FLAG = False Then
                SELPATHNO_QR(0) = SELREQNO_STFRS("REQNNO")    '                SELRREQUISITN_RSALL("REQNNO")
                Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                SELPATHNO_FLAG = True
            Else
                SELPATHNO_QR(0) = SELREQNO_STFRS("REQNNO") ' SELRREQUISITN_RSALL("REQNNO")
                SELPATHNO_RS.MoreResults
                SELPATHNO_RS.Requery
            End If
            If SELMA101FLAG = False Then
                 SELMA101_QR(0) = SELREQNO_STFRS("SER_CODE")
                 Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                 SELMA101FLAG = True
            Else
                 SELMA101_QR(0) = SELREQNO_STFRS("SER_CODE")
                 SELMA101_RS.MoreResults
                 SELMA101_RS.Requery
            End If
            If SELMA101_RS.RowCount > 0 Then
                If Trim(SELREQNO_STFRS("REQNSTAT")) = "C" Then
                    MSFlexGrid1.AddItem i & vbTab & SELREQNO_STFRS("MODALITYCD") & vbTab & req_chglen(SELREQNO_STFRS("REQNNO"))
                Else
                    MSFlexGrid1.AddItem i & vbTab & SELREQNO_STFRS("MODALITYCD") & vbTab & req_chglen(SELREQNO_STFRS("REQNNO"))
                End If
            i = i + 1
            End If
            SELREQNO_STFRS.MoveNext
        Wend
        chk_collected.Visible = True
    Else
        MsgBox "There are No Pending Requisition No."
    End If
End If


End Sub
Public Sub ret_pat_data()
Dim i As Integer
Dim j As Integer
'On Error GoTo MYERROR:
If NEWFLAG = True Then
    If Trim(cmb_location.Text) = "GEN.OPD" Or Trim(cmb_location.Text) = "PVT.OPD" Or (OPT_CASENO.Value = True) Then
        If selrreqnflagall = False Then
            selrreqn_opqr(0) = msk_caseno.FormattedText
            selrreqn_opqr(1) = msk_fromdt.FormattedText
            selrreqn_opqr(2) = msk_todt.FormattedText
            selrreqn_opqr(3) = "P"
            selrreqn_opqr(4) = "P"
            Set selrreqn_oprs = selrreqn_opqr.OpenResultset(1, 3)
        Else
            selrreqn_opqr(0) = msk_caseno.FormattedText
            selrreqn_opqr(1) = msk_fromdt.FormattedText
            selrreqn_opqr(2) = msk_todt.FormattedText
            selrreqn_opqr(3) = "P"
            selrreqn_opqr(4) = "P"
            selrreqn_oprs.MoreResults
            selrreqn_oprs.Requery
        End If
        
        '*********** Shraddha Pandit 16/07/09 *************
        sel_ssrreqn_qr(0) = msk_caseno.FormattedText
        sel_ssrreqn_qr(1) = msk_fromdt.FormattedText
        sel_ssrreqn_qr(2) = msk_todt.FormattedText
        Set sel_ssrreqn_rs = sel_ssrreqn_qr.OpenResultset(1, 3)
        '**************************************************
        
        NEWFLAG = False
                            
        If selrreqn_oprs.RowCount > 0 Then
        
            txt_rowcount.Text = "Total Number of Requisition for Which Sample have to Collect    " & selrreqn_oprs.RowCount
            i = 1
            MSFlexGrid1.Rows = 1
            
            While Not selrreqn_oprs.EOF
                If Trim(msk_caseno.FormattedText) = selrreqn_oprs("CASENO") Then
                    If SELPATHNO_FLAG = False Then
                        SELPATHNO_QR(0) = selrreqn_oprs("REQNNO")
                        Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                        SELPATHNO_FLAG = True
                    Else
                        SELPATHNO_QR(0) = selrreqn_oprs("REQNNO")
                        SELPATHNO_RS.MoreResults
                        SELPATHNO_RS.Requery
                    End If
                    
                    
                    
                    
                    
                    If Trim(selrreqn_oprs("ser_code")) <> "8888" Then
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = selrreqn_oprs("SER_CODE")
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = selrreqn_oprs("SER_CODE")
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            
                            MSFlexGrid1.AddItem i & vbTab & selrreqn_oprs("CASENO") & vbTab & req_chglen(selrreqn_oprs("REQNNO")) & "/" & Trim(selrreqn_oprs("MEMONO"))
                            '********* Shraddha 06/07/09 ********
                            MSFlexGrid1.TextMatrix(i, 3) = "Print Barcode"
                            '************************************
                            '***** Shraddha 18/11/2011 *****
'                            str = "SELECT distinct SAMPLECODE.SAMPLECD FROM DISLIB.SAMPLECODE,DISLIB.DMULPROCE WHERE DMULPROCE.PRO_CODE=SAMPLECODE.SERCODE AND DMULPROCE.REQNO=?"
'                            Set SELLABNO_QR = rdb.CreateQuery("", str)
'
'                            SELLABNO_QR(0) = Trim(selrreqn_oprs("REQNNO"))
'                            'SELLABNO_QR.Execute
'                            Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
'                            DESCRIP = ""
'                            While Not SELLABNO_RS.EOF
'                                smplcd = SELLABNO_RS(0)
'                                deptdesc
'
'                                DPT = Mid(Trim(selrreqn_oprs("REQNNO")), 1, 3)
'
'                                SELDEPTDES_QR(0) = DPT
'                                Set SELDEPTDES_RS = SELDEPTDES_QR.OpenResultset(1, 3)
'
'                                'DESCRIP = ""
'                                'DESCRIP = Mid(Trim(SELDEPTDES_RS("DEPT_NAME")), 1, 4) & "/" & colorcd
'                                DESCRIP = DESCRIP & " " & Mid(Trim(SELDEPTDES_RS("DEPT_NAME")), 1, 4) & "/" & colorcd
'
'                                SELLABNO_RS.MoveNext
'                            Wend
                    '***************************************
                            str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
                            Set SEL_DMUL_QR = rdb.CreateQuery("", str)

                            SEL_DMUL_QR(0) = Trim(selrreqn_oprs("REQNNO"))
                            'SELLABNO_QR.Execute
                            Set SEL_DMUL_RS = SEL_DMUL_QR.OpenResultset(1, 3)
                            TESTDESC = ""
                            While Not SEL_DMUL_RS.EOF
                                                          
                                str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
                                Set SEL_MA101_QR = rdb.CreateQuery("", str)
                                                          
                                SEL_MA101_QR(0) = Trim(SEL_DMUL_RS("PRO_CODE"))
                                Set SEL_MA101_RS = SEL_MA101_QR.OpenResultset(1, 3)

                                'DESCRIP = ""
                                'DESCRIP = Mid(Trim(SELDEPTDES_RS("DEPT_NAME")), 1, 4) & "/" & colorcd
                                TESTDESC = TESTDESC & "| " & Trim(SEL_MA101_RS("SER_DET"))
                                SEL_DMUL_RS.MoveNext
                            Wend
                    
                            MSFlexGrid1.TextMatrix(i, 4) = TESTDESC
                            '*******************************
                            
                            i = i + 1
                        End If
                    Else
                        
                        MSFlexGrid1.AddItem i & vbTab & selrreqn_oprs("CASENO") & vbTab & req_chglen(selrreqn_oprs("REQNNO"))
                        '********* Shraddha 06/07/09 ********
                            MSFlexGrid1.TextMatrix(i, 3) = "Print Barcode"
                        '************************************
                        '***** Shraddha 18/11/2011 *****
                        str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
                            Set SEL_DMUL_QR = rdb.CreateQuery("", str)

                            SEL_DMUL_QR(0) = Trim(selrreqn_oprs("REQNNO"))
                            'SELLABNO_QR.Execute
                            Set SEL_DMUL_RS = SEL_DMUL_QR.OpenResultset(1, 3)
                            TESTDESC = ""
                            While Not SEL_DMUL_RS.EOF
                                                          
                                str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
                                Set SEL_MA101_QR = rdb.CreateQuery("", str)
                                                          
                                SEL_MA101_QR(0) = Trim(SEL_DMUL_RS("PRO_CODE"))
                                Set SEL_MA101_RS = SEL_MA101_QR.OpenResultset(1, 3)

                                'DESCRIP = ""
                                'DESCRIP = Mid(Trim(SELDEPTDES_RS("DEPT_NAME")), 1, 4) & "/" & colorcd
                                    If SEL_MA101_RS.EOF = False Then
                                TESTDESC = TESTDESC & "| " & Trim(SEL_MA101_RS("SER_DET"))
                                End If
                                SEL_DMUL_RS.MoveNext
                            Wend
                    
                            MSFlexGrid1.TextMatrix(i, 4) = TESTDESC
                        '*******************************
                        
                        
                        i = i + 1
                    End If
                    
                   
                    
                    
                    selrreqn_oprs.MoveNext
                Else
                    selrreqn_oprs.MoveNext
                End If
            Wend
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 3
            chk_collected.Width = MSFlexGrid1.CellWidth
            chk_collected.Height = MSFlexGrid1.CellHeight
            chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_collected.Value = 0
            chk_collected.Visible = True
        End If
        
        '************* Shraddha Pandit 16/07/09 *****************
        If sel_ssrreqn_rs.RowCount > 0 Then
            'txt_rowcount.Text = "Total Number of Requisition for Which Sample have to Collect    " & selrreqn_oprs.RowCount
            j = i
            'following line commented by shweta 13-03-2012
            'MSFlexGrid1.Rows = 1
            'and condition added by shweta 13-03-2012
            If MSFlexGrid1.Rows = 0 Then
            MSFlexGrid1.Rows = 1
            Else
            On Error GoTo aa11
               If (Trim(MSFlexGrid1.TextMatrix(1, 2)) = "") And (Trim(MSFlexGrid1.TextMatrix(1, 3)) = "") Then
               MSFlexGrid1.Rows = 1
               End If
aa11:
            End If
            
         
            While Not sel_ssrreqn_rs.EOF
                If Trim(msk_caseno.FormattedText) = sel_ssrreqn_rs("CASENO") Then
                    If SELPATHNO_FLAG = False Then
                        SELPATHNO_QR(0) = sel_ssrreqn_rs("REQNNO")
                        Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                        SELPATHNO_FLAG = True
                    Else
                        SELPATHNO_QR(0) = sel_ssrreqn_rs("REQNNO")
                        SELPATHNO_RS.MoreResults
                        SELPATHNO_RS.Requery
                    End If

                    If Trim(sel_ssrreqn_rs("ser_code")) <> "8888" Then
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = sel_ssrreqn_rs("SER_CODE")
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = sel_ssrreqn_rs("SER_CODE")
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        
                        
                        
                        If SELMA101_RS.RowCount > 0 Then

                            'MSFlexGrid1.AddItem j & vbTab & sel_ssrreqn_rs("CASENO") & vbTab & req_chglen(sel_ssrreqn_rs("REQNNO"))
                            'Commented and added by Shraddha on 05/12/2011
                            MSFlexGrid1.AddItem j & vbTab & sel_ssrreqn_rs("CASENO") & vbTab & req_chglen(sel_ssrreqn_rs("REQNNO")) & Trim(sel_ssrreqn_rs("MEMONO"))
                            '********* Shraddha 06/07/09 ********
                            MSFlexGrid1.TextMatrix(j, 3) = "Print Barcode"
                            '************************************
                            j = j + 1
                        End If
                    Else

                        'MSFlexGrid1.AddItem j & vbTab & sel_ssrreqn_rs("CASENO") & vbTab & req_chglen(sel_ssrreqn_rs("REQNNO"))
                        'Commented and added by Shraddha on 05/12/2011
                        MSFlexGrid1.AddItem j & vbTab & sel_ssrreqn_rs("CASENO") & vbTab & req_chglen(sel_ssrreqn_rs("REQNNO")) & Trim(sel_ssrreqn_rs("MEMONO"))
                        '********* Shraddha 06/07/09 ********
                        MSFlexGrid1.TextMatrix(j, 3) = "Print Barcode"
                        '************************************
                        
                
                        j = j + 1
                    End If

                    sel_ssrreqn_rs.MoveNext
                Else
                    selrreqn_oprs.MoveNext
                End If
            Wend
            '***** Below code uncommented by Shraddha on 10/02/10 *****
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 3
            chk_collected.Width = MSFlexGrid1.CellWidth
            chk_collected.Height = MSFlexGrid1.CellHeight
            chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_collected.Value = 0
            chk_collected.Visible = True
            '***********************************************
        End If
        '********************************************************
        
    End If
End If
    
If NEWFLAG = True Then
    NEWFLAG = False
    'SELREQNO_QR(0) = msk_caseno.FormattedText
    SELREQNO_QR(0) = "P"
    SELREQNO_QR(1) = "P"
    SELREQNO_QR(2) = GET_REVDATE(msk_fromdt.FormattedText)
    SELREQNO_QR(3) = GET_REVDATE(msk_todt.FormattedText)
    'SELREQNO_QR(4) = Trim(cmb_location.Text)
    If SELREQNOFLAG = False Then
        Set SELREQNO_RS = SELREQNO_QR.OpenResultset(1, 3)
        SELREQNOFLAG = True
    Else
        SELREQNO_RS.MoreResults
        SELREQNO_RS.Requery
    End If
    If SELREQNO_RS.RowCount > 0 Then
        txt_rowcount.Text = "Total Number of Requisition for Which Sample have to Collect    " & SELREQNO_RS.RowCount
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELREQNO_RS.EOF
            If Trim(msk_caseno.FormattedText) = SELREQNO_RS("CASENO") Then
                If SELPATHNO_FLAG = False Then
                    SELPATHNO_QR(0) = SELREQNO_RS("REQNNO")
                    Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                    SELPATHNO_FLAG = True
                Else
                    SELPATHNO_QR(0) = SELREQNO_RS("REQNNO")
                    SELPATHNO_RS.MoreResults
                    SELPATHNO_RS.Requery
                End If
                If Trim(SELREQNO_RS("ser_code")) <> "8888" Then
                    If SELMA101FLAG = False Then
                        SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                        SELMA101FLAG = True
                    Else
                        SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
                        SELMA101_RS.MoreResults
                        SELMA101_RS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        MSFlexGrid1.AddItem i & vbTab & SELREQNO_RS("CASENO") & vbTab & req_chglen(SELREQNO_RS("REQNNO"))
                        i = i + 1
                    End If
                Else
                    MSFlexGrid1.AddItem i & vbTab & SELREQNO_RS("CASENO") & vbTab & req_chglen(SELREQNO_RS("REQNNO"))
                    i = i + 1
                End If
                SELREQNO_RS.MoveNext
            Else
                SELREQNO_RS.MoveNext
            End If
        Wend
        MSFlexGrid1.row = 1
        MSFlexGrid1.col = 3
        chk_collected.Width = MSFlexGrid1.CellWidth
        chk_collected.Height = MSFlexGrid1.CellHeight
        chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
        chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        chk_collected.Value = 0
        chk_collected.Visible = True
    End If
End If

If viewflag = True Then
    viewflag = FLASE
    SELREQNO_QRVIEW(0) = msk_caseno.FormattedText
    SELREQNO_QRVIEW(1) = "C" 'Shraddha Pandit 28/05/2009
    SELREQNO_QRVIEW(2) = "T"
    SELREQNO_QRVIEW(3) = GET_REVDATE(msk_fromdt.FormattedText)
    SELREQNO_QRVIEW(4) = GET_REVDATE(msk_todt.FormattedText)
    'SELREQNO_QRVIEW(5) = Trim(cmb_location.Text) 'Shraddha Pandit 02/02/09
    'SELREQNO_QRVIEW(6) = Trim(cmb_location.Text) 'Shraddha Pandit 02/02/09
    If SELREQNOFLAGVIEW = False Then
        Set SELREQNO_RSVIEW = SELREQNO_QRVIEW.OpenResultset(1, 3)
        SELREQNOFLAGVIEW = True
    Else
        SELREQNO_RSVIEW.MoreResults
        SELREQNO_RSVIEW.Requery
    End If
    If SELREQNO_RSVIEW.RowCount > 0 Then
        txt_rowcount.Text = "Total Number of Requisition for Which Sample have Collected    " & SELREQNO_RSVIEW.RowCount
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELREQNO_RSVIEW.EOF
            If SELPATHNO_FLAG = False Then
                SELPATHNO_QR(0) = SELREQNO_RSVIEW("REQNNO")
                Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                SELPATHNO_FLAG = True
            Else
                SELPATHNO_QR(0) = SELREQNO_RSVIEW("REQNNO")
                SELPATHNO_RS.MoreResults
                SELPATHNO_RS.Requery
            End If
            If SELREQNO_RSVIEW("SER_CODE") <> "8888" Then
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = SELREQNO_RSVIEW("SER_CODE")
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = SELREQNO_RSVIEW("SER_CODE")
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    If Trim(SELREQNO_RSVIEW("REQNSTAT")) = "C" Then
                        MSFlexGrid1.AddItem i & vbTab & SELREQNO_RSVIEW("CASENO") & vbTab & req_chglen(SELREQNO_RSVIEW("REQNNO"))
                    Else
                        MSFlexGrid1.AddItem i & vbTab & SELREQNO_RSVIEW("CASENO") & vbTab & req_chglen(SELREQNO_RSVIEW("REQNNO"))
                    End If
                    i = i + 1
                End If
            Else
                If Trim(SELREQNO_RSVIEW("REQNSTAT")) = "C" Then
                    MSFlexGrid1.AddItem i & vbTab & SELREQNO_RSVIEW("CASENO") & vbTab & req_chglen(SELREQNO_RSVIEW("REQNNO"))
                Else
                    MSFlexGrid1.AddItem i & vbTab & SELREQNO_RSVIEW("CASENO") & vbTab & req_chglen(SELREQNO_RSVIEW("REQNNO"))
                End If
                i = i + 1
            End If
            SELREQNO_RSVIEW.MoveNext
        Wend
        MSFlexGrid1.row = 1
        MSFlexGrid1.col = 3
        chk_collected.Width = MSFlexGrid1.CellWidth
        chk_collected.Height = MSFlexGrid1.CellHeight
        chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
        chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        chk_collected.Value = 0
        chk_collected.Visible = True
    Else
        MsgBox "There are No Pending Requisition No."
    End If
End If
Exit Sub
MYERROR:

MsgBox Err.Description
End Sub
Private Sub Opt_all_Click()
    If OPT_ALL.Value = True Then
        msk_fromdt.Enabled = True
        msk_todt.Enabled = True
        Opt_patient.Enabled = True
        opt_staff.Enabled = True
        OPT_CASENO.Enabled = True
        msk_caseno.Enabled = False
        'msk_fromdt.Text = DATE11
        msk_fromdt.Text = P_LAST_WORKING_DAY
        msk_fromdt.Text = mydate
        msk_todt.Text = DATE11
        'Msk_fromdt.SetFocus
        casenoflag = False
        Opt_patient.Value = False
        opt_staff.Value = False
        lbl_name.Visible = False
        txt_name.Visible = False
        ''' done by albert ///****
        Frame3.Visible = False
        Label4.Visible = False
        msk_caseno.Visible = False
        lbl_name.Visible = False
        txt_name.Visible = False
        Label3.Visible = True
        cmb_dept.Visible = True
        Label1.Visible = True
        cmb_modality.Visible = True
        cmb_modality.Enabled = True
        cmb_dept.Visible = True
        cmb_dept.Enabled = True
        '''up to/****
    End If

End Sub

Private Sub OPT_CASENO_Click()
  'Opt_patient.Value = True
   lbl_name.Visible = True
   txt_name.Visible = True
   msk_caseno.Enabled = True
''' done by albert ///****
        Frame3.Visible = True
        Label4.Visible = True
        msk_caseno.Visible = True
        lbl_name.Visible = True
        txt_name.Visible = True
        Label3.Visible = False
        cmb_dept.Visible = False
        Label1.Visible = False
        cmb_modality.Visible = False
        '''up to/****
  'msk_caseno.SetFocus
End Sub
Private Sub OPT_CASENO_LostFocus()
Opt_patient.Value = True
End Sub

Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
    OPT_CASENO.Caption = "Case No."
   ' msk_caseno.Mask = "??>/#####"
    OPT_CASENO.Enabled = True
    'Opt_all.Enabled = False
    msk_fromdt.Enabled = False

    msk_todt.Enabled = True
    msk_caseno.Enabled = True
    allflag = False
    Label4.Caption = "Case No."
    msk_caseno.SetFocus
    'msk_caseno.SetFocus
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    OPT_CASENO.Caption = "CC No."
    msk_caseno.Mask = "######"
    OPT_CASENO.Enabled = True
    OPT_ALL.Enabled = False
    msk_fromdt.Enabled = False
    msk_todt.Enabled = True
    OPT_CASENO.Value = True
    msk_caseno.Enabled = True
    Label4.Caption = "CC No."
    msk_caseno.SetFocus
End If
End Sub

Public Sub STAFF()
    If SELDREQSTAFFDFLAG = False Then
        SELDREQSTAFFD_QR(0) = msk_caseno.FormattedText
        Set SELDREQSTAFFD_RS = SELDREQSTAFFD_QR.OpenResultset(1, 3)
        SELDREQSTAFFDFLAG = True
    Else
        SELDREQSTAFFD_QR(0) = msk_caseno.FormattedText
        SELDREQSTAFFD_RS.MoreResults
        SELDREQSTAFFD_RS.Requery
    End If
    If SELDREQSTAFFD_RS.RowCount > 0 Then
        txt_name = SELDREQSTAFFD_RS("PNAME")
        msk_fromdt.Enabled = True
    msk_fromdt.SetFocus
    End If
End Sub
Private Sub ALLCASE()
If SELRREQUISITN_FLAGCASE = False Then
    SELRREQUISITN_QRCASE(0) = msk_caseno.FormattedText
    SELRREQUISITN_QRCASE(1) = "P"
    SELRREQUISITN_QRCASE(2) = "P"
    Set SELRREQUISITN_RSCASE = SELRREQUISITN_QRCASE.OpenResultset(1, 3)
Else
    SELRREQUISITN_QRCASE(0) = msk_caseno.FormattedText
    SELRREQUISITN_QRCASE(1) = "P"
    SELRREQUISITN_QRCASE(2) = "P"
    SELRREQUISITN_RSCASE.MoreResults
    SELRREQUISITN_RSCASE.Requery
End If
If SELRREQUISITN_RSCASE.RowCount > 0 Then
txt_rowcount.Text = "Total Number of Requisition for Which Sample have to Collect    " & SELRREQUISITN_RSCASE.RowCount
    i = 1
    MSFlexGrid1.Rows = 1
    While Not SELRREQUISITN_RSCASE.EOF
        If SELPATHNO_FLAG = False Then
            SELPATHNO_QR(0) = SELRREQUISITN_RSCASE("REQNNO")
            Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
            SELPATHNO_FLAG = True
        Else
            SELPATHNO_QR(0) = SELRREQUISITN_RSCASE("REQNNO")
            SELPATHNO_RS.MoreResults
            SELPATHNO_RS.Requery
        End If
        If Trim(SELRREQUISITN_RSCASE("ser_code")) <> "8888" Then
             If SELMA101FLAG = False Then
                 SELMA101_QR(0) = SELRREQUISITN_RSCASE("SER_CODE")
                 Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                 SELMA101FLAG = True
             Else
                 SELMA101_QR(0) = SELRREQUISITN_RSCASE("SER_CODE")
                 SELMA101_RS.MoreResults
                 SELMA101_RS.Requery
            End If
            If SELMA101_RS.RowCount > 0 Then
                MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RSCASE("CASENO") & vbTab & req_chglen(SELRREQUISITN_RSCASE("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELRREQUISITN_RSCASE("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                i = i + 1
            End If
        Else
            MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RSCASE("CASENO") & vbTab & req_chglen(SELRREQUISITN_RSCASE("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELRREQUISITN_RSCASE("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
            i = i + 1
        End If
        SELRREQUISITN_RSCASE.MoveNext
    Wend
    MSFlexGrid1.row = 1
'    MSFlexGrid1.col = 5
    'Commented and added by Shraddha on 21/11/2011
    MSFlexGrid1.col = 3
    chk_collected.Width = MSFlexGrid1.CellWidth
    chk_collected.Height = MSFlexGrid1.CellHeight
    chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    chk_collected.Value = 0
    chk_collected.Visible = True
End If
End Sub

Public Sub RETRIVE_DATA()
Dim i As Integer
On Error GoTo MYERROR:
If NEWFLAG = True Then
    If Trim(cmb_location.Text) = "GEN.OPD" Or Trim(cmb_location.Text) = "PVT.OPD" Or (OPT_CASENO.Value = True) Then
        If selrreqnflagall = False Then
                selrreqn_opqr(0) = msk_caseno.FormattedText
                selrreqn_opqr(1) = msk_fromdt.FormattedText
                selrreqn_opqr(2) = msk_todt.FormattedText
                selrreqn_opqr(3) = "P"
                selrreqn_opqr(4) = "C"
                Set selrreqn_oprs = selrreqn_opqr.OpenResultset(1, 3)
            Else
                selrreqn_opqr(0) = msk_caseno.FormattedText
                selrreqn_opqr(1) = msk_fromdt.FormattedText
                selrreqn_opqr(2) = msk_todt.FormattedText
                selrreqn_opqr(3) = "P"
                selrreqn_opqr(4) = "C"
                selrreqn_oprs.MoreResults
                selrreqn_oprs.Requery
            End If
            NEWFLAG = False
        If selrreqn_oprs.RowCount > 0 Then
          txt_rowcount.Text = "Total Number of Requisition for Which Sample have to Collect    " & selrreqn_oprs.RowCount
          i = 1
          MSFlexGrid1.Rows = 1
          While Not selrreqn_oprs.EOF
            If Trim(msk_caseno.FormattedText) = selrreqn_oprs("CASENO") Then
              If SELPATHNO_FLAG = False Then
                  SELPATHNO_QR(0) = selrreqn_oprs("REQNNO")
                  Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                  SELPATHNO_FLAG = True
              Else
                  SELPATHNO_QR(0) = selrreqn_oprs("REQNNO")
                  SELPATHNO_RS.MoreResults
                  SELPATHNO_RS.Requery
              End If
              If Trim(selrreqn_oprs("ser_code")) <> "8888" Then
                  If SELMA101FLAG = False Then
                      SELMA101_QR(0) = selrreqn_oprs("SER_CODE")
                      Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                      SELMA101FLAG = True
                  Else
                      SELMA101_QR(0) = selrreqn_oprs("SER_CODE")
                      SELMA101_RS.MoreResults
                      SELMA101_RS.Requery
                  End If
                  If SELMA101_RS.RowCount > 0 Then
                      MSFlexGrid1.AddItem i & vbTab & selrreqn_oprs("CASENO") & vbTab & req_chglen(selrreqn_oprs("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & selrreqn_oprs("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                      i = i + 1
                  End If
              Else
                  MSFlexGrid1.AddItem i & vbTab & selrreqn_oprs("CASENO") & vbTab & req_chglen(selrreqn_oprs("REQNNO")) & vbTab & "M.SERVICES" & vbTab & selrreqn_oprs("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                  i = i + 1
              End If
              selrreqn_oprs.MoveNext
            Else
              selrreqn_oprs.MoveNext
            End If
          Wend
          MSFlexGrid1.row = 1
          MSFlexGrid1.col = 3
          chk_collected.Width = MSFlexGrid1.CellWidth
          chk_collected.Height = MSFlexGrid1.CellHeight
          chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
          chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
          chk_collected.Value = 0
          chk_collected.Visible = True
      End If
    End If
End If

If NEWFLAG = True Then
    NEWFLAG = False
    'SELREQNO_QR(0) = msk_caseno.FormattedText
    SELREQNO_QR(0) = "P"
    SELREQNO_QR(1) = "P"
    SELREQNO_QR(2) = GET_REVDATE(msk_fromdt.FormattedText)
    SELREQNO_QR(3) = GET_REVDATE(msk_todt.FormattedText)
    'SELREQNO_QR(4) = Trim(cmb_location.Text)
    If SELREQNOFLAG = False Then
        Set SELREQNO_RS = SELREQNO_QR.OpenResultset(1, 3)
        SELREQNOFLAG = True
    Else
        SELREQNO_RS.MoreResults
        SELREQNO_RS.Requery
    End If
        If SELREQNO_RS.RowCount > 0 Then

                txt_rowcount.Text = "Total Number of Requisition for Which Sample have to Collect    " & SELREQNO_RS.RowCount
                i = 1
                MSFlexGrid1.Rows = 1
                While Not SELREQNO_RS.EOF
                  If Trim(msk_caseno.FormattedText) = SELREQNO_RS("CASENO") Then
                    If SELPATHNO_FLAG = False Then
                        SELPATHNO_QR(0) = SELREQNO_RS("REQNNO")
                        Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                        SELPATHNO_FLAG = True
                    Else
                        SELPATHNO_QR(0) = SELREQNO_RS("REQNNO")
                        SELPATHNO_RS.MoreResults
                        SELPATHNO_RS.Requery
                    End If
                    If Trim(SELREQNO_RS("ser_code")) <> "8888" Then
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = SELREQNO_RS("SER_CODE")
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            MSFlexGrid1.AddItem i & vbTab & SELREQNO_RS("CASENO") & vbTab & req_chglen(SELREQNO_RS("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELREQNO_RS("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                            i = i + 1
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & SELREQNO_RS("CASENO") & vbTab & req_chglen(SELREQNO_RS("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELREQNO_RS("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                        i = i + 1
                    End If
                    SELREQNO_RS.MoveNext

                    Else
                        SELREQNO_RS.MoveNext
                    End If

                Wend
                MSFlexGrid1.row = 1
                MSFlexGrid1.col = 3
                chk_collected.Width = MSFlexGrid1.CellWidth
                chk_collected.Height = MSFlexGrid1.CellHeight
                chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                chk_collected.Value = 0
                chk_collected.Visible = True
        End If
    End If
If viewflag = True Then
    viewflag = FLASE
    SELREQNO_QRVIEW(0) = msk_caseno.FormattedText
    SELREQNO_QRVIEW(1) = "C"
    SELREQNO_QRVIEW(2) = "C"
    SELREQNO_QRVIEW(3) = GET_REVDATE(msk_fromdt.FormattedText)
    SELREQNO_QRVIEW(4) = GET_REVDATE(msk_todt.FormattedText)
    'SELREQNO_QRVIEW(5) = Trim(cmb_location.Text) 'Shraddha Pandit 02/02/09
    'SELREQNO_QRVIEW(6) = Trim(cmb_location.Text) 'Shraddha Pandit 02/02/09
    If SELREQNOFLAGVIEW = False Then
        Set SELREQNO_RSVIEW = SELREQNO_QRVIEW.OpenResultset(1, 3)
        SELREQNOFLAGVIEW = True
    Else
        SELREQNO_RSVIEW.MoreResults
        SELREQNO_RSVIEW.Requery
    End If
    If SELREQNO_RSVIEW.RowCount > 0 Then
        txt_rowcount.Text = "Total Number of Requisition for Which Sample have Collected    " & SELREQNO_RSVIEW.RowCount
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELREQNO_RSVIEW.EOF
            If SELPATHNO_FLAG = False Then
                SELPATHNO_QR(0) = SELREQNO_RSVIEW("REQNNO")
                Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                SELPATHNO_FLAG = True
            Else
                SELPATHNO_QR(0) = SELREQNO_RSVIEW("REQNNO")
                SELPATHNO_RS.MoreResults
                SELPATHNO_RS.Requery
            End If
            If SELREQNO_RSVIEW("SER_CODE") <> "8888" Then
                 If SELMA101FLAG = False Then
                     SELMA101_QR(0) = SELREQNO_RSVIEW("SER_CODE")
                     Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                     SELMA101FLAG = True
                 Else
                     SELMA101_QR(0) = SELREQNO_RSVIEW("SER_CODE")
                     SELMA101_RS.MoreResults
                     SELMA101_RS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    If Trim(SELREQNO_RSVIEW("REQNSTAT")) = "C" Then
                        MSFlexGrid1.AddItem i & vbTab & SELREQNO_RSVIEW("MODALITYCD") & vbTab & req_chglen(SELREQNO_RSVIEW("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELREQNO_RSVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                    Else
                        MSFlexGrid1.AddItem i & vbTab & SELREQNO_RSVIEW("MODALITYCD") & vbTab & req_chglen(SELREQNO_RSVIEW("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELREQNO_RSVIEW("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                    End If
                i = i + 1
                End If
            Else
                If Trim(SELREQNO_RSVIEW("REQNSTAT")) = "C" Then
                    MSFlexGrid1.AddItem i & vbTab & SELREQNO_RSVIEW("MODALITYCD") & vbTab & req_chglen(SELREQNO_RSVIEW("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELREQNO_RSVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                Else
                    MSFlexGrid1.AddItem i & vbTab & SELREQNO_RSVIEW("MODALITYCD") & vbTab & req_chglen(SELREQNO_RSVIEW("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELREQNO_RSVIEW("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                End If
                i = i + 1
            End If
            SELREQNO_RSVIEW.MoveNext
        Wend
        chk_collected.Visible = False
    Else
        MsgBox "There are No Pending Requisition No."
    End If
End If
Exit Sub
MYERROR:
MsgBox Err.Description
End Sub



Public Sub RETRIV_STAFF_DATE()


If NEWFLAG = True Then
    If (Trim(cmb_location.Text) = "STAFF") Or (OPT_CASENO.Value = True) Then
        NEWFLAG = False
        SELREQNO_STFQR(0) = msk_caseno.FormattedText
        SELREQNO_STFQR(1) = "P"
        SELREQNO_STFQR(2) = "P"
        SELREQNO_STFQR(3) = GET_REVDATE(msk_fromdt.FormattedText)
        SELREQNO_STFQR(4) = GET_REVDATE(msk_todt.FormattedText)
        'SELREQNO_STFQR(4) = Trim(cmb_location.Text)

            If SELREQNOSTFFLAG = False Then
                Set SELREQNO_STFRS = SELREQNO_STFQR.OpenResultset(1, 3)
                SELREQNOSTFFLAG = True
            Else
                SELREQNO_STFRS.MoreResults
                SELREQNO_STFRS.Requery
            End If

            If SELREQNO_STFRS.RowCount > 0 Then
                i = 1
                MSFlexGrid1.Rows = 1
                While Not SELREQNO_STFRS.EOF
                    'If Trim(msk_caseno.FormattedText) = SELREQNO_STFQR(0) Then
                        If SELPATHNO_FLAG = False Then
                            SELPATHNO_QR(0) = SELREQNO_STFRS("REQNNO")
                            Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                            SELPATHNO_FLAG = True
                        Else
                            SELPATHNO_QR(0) = SELREQNO_STFRS("REQNNO")
                            SELPATHNO_RS.MoreResults
                            SELPATHNO_RS.Requery
                        End If
                    If Trim(SELREQNO_STFRS("ser_code")) <> "8888" Then
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = SELREQNO_STFRS("SER_CODE")
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = SELREQNO_STFRS("SER_CODE")
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            MSFlexGrid1.AddItem i & vbTab & SELREQNO_STFRS("CASENO") & vbTab & req_chglen(SELREQNO_STFRS("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELREQNO_STFRS("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                            i = i + 1
                        End If
                     Else
                            MSFlexGrid1.AddItem i & vbTab & SELREQNO_STFRS("CASENO") & vbTab & req_chglen(SELREQNO_STFRS("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELREQNO_STFRS("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                            i = i + 1
                     End If
                        SELREQNO_STFRS.MoveNext
                    'Else
                       ' SELREQNO_STFRS.MoveNext
                    'End If
                Wend
                    MSFlexGrid1.row = 1
                    'MSFlexGrid1.col = 5
                    'Commented and added by Shraddha on 21/11/2011
                    MSFlexGrid1.col = 3
                    chk_collected.Width = MSFlexGrid1.CellWidth
                    chk_collected.Height = MSFlexGrid1.CellHeight
                    chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                    chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                    chk_collected.Value = 0
                    chk_collected.Visible = True
            End If
    Else
            If Trim(cmb_location.Text) = "STAFF" And Trim(cmb_dept.Text) = "ALL" Then
                NEWFLAG = False
                'SELREQNO_QR(0) = msk_caseno.FormattedText
                SELREQNO_STFQRALL(0) = "P"
                SELREQNO_STFQRALL(1) = "P"
                SELREQNO_STFQRALL(2) = GET_REVDATE(msk_fromdt.FormattedText)
                SELREQNO_STFQRALL(3) = GET_REVDATE(msk_todt.FormattedText)
                'SELREQNO_STFQR(4) = Trim(cmb_location.Text)

                If SELREQNOSTFFLAGALL = False Then
                    Set SELREQNO_STFRSALL = SELREQNO_STFQRALL.OpenResultset(1, 3)
                    SELREQNOSTFFLAGALL = True
                Else
                    SELREQNO_STFRSALL.MoreResults
                    SELREQNO_STFRSALL.Requery
                End If

                If SELREQNO_STFRSALL.RowCount > 0 Then
                    i = 1
                    MSFlexGrid1.Rows = 1
                    While Not SELREQNO_STFRSALL.EOF

                        If SELPATHNO_FLAG = False Then
                            SELPATHNO_QR(0) = SELREQNO_STFRSALL("REQNNO")
                            Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                            SELPATHNO_FLAG = True
                        Else
                            SELPATHNO_QR(0) = SELREQNO_STFRSALL("REQNNO")
                            SELPATHNO_RS.MoreResults
                            SELPATHNO_RS.Requery
                        End If
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = SELREQNO_STFRSALL("SER_CODE")
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = SELREQNO_STFRSALL("SER_CODE")
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            MSFlexGrid1.AddItem i & vbTab & SELREQNO_STFRSALL("CASENO") & vbTab & req_chglen(SELREQNO_STFRSALL("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELREQNO_STFRSALL("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                            i = i + 1
                        End If
                            SELREQNO_STFRSALL.MoveNext

                    Wend
                        MSFlexGrid1.row = 1
                        'MSFlexGrid1.col = 5
                        'Commented and added by Shraddha on 21/11/2011
                        MSFlexGrid1.col = 3
                        chk_collected.Width = MSFlexGrid1.CellWidth
                        chk_collected.Height = MSFlexGrid1.CellHeight
                        chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                        chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                        chk_collected.Value = 0
                        chk_collected.Visible = True
                End If
            End If
    End If

    If (Trim(cmb_dept.Text) <> "ALL") And (cmb_location.Text = "STAFF") Then
        If Trim(cmb_modality.Text) = "" Then
            NEWSTFALLMODALITY
        Else
            If SELRREQUISITN_STFFLAG = False Then
                SELRREQUISITN_STFQR(0) = msk_fromdt.FormattedText
                SELRREQUISITN_STFQR(1) = msk_todt.FormattedText
                SELRREQUISITN_STFQR(2) = "P"
                SELRREQUISITN_STFQR(3) = "P"
                SELRREQUISITN_STFQR(4) = dept
                i = InStr(cmb_modality.Text, "$")
                SELRREQUISITN_STFQR(5) = Mid(cmb_modality, i + 1, 2)
                'SELRREQUISITN_QR(6) = Trim(cmb_location.Text)
                Set SELRREQUISITN_STFRS = SELRREQUISITN_STFQR.OpenResultset(1, 3)
            Else
                SELRREQUISITN_STFQR(0) = msk_fromdt.FormattedText
                SELRREQUISITN_STFQR(1) = msk_todt.FormattedText
                SELRREQUISITN_STFQR(2) = "P"
                SELRREQUISITN_STFQR(3) = "P"
                SELRREQUISITN_STFQR(4) = dept
                i = InStr(cmb_modality.Text, "$")
                SELRREQUISITN_STFQR(5) = Mid(cmb_modality, i + 1, 2)
                'SELRREQUISITN_QR(6) = Trim(cmb_location.Text)
                SELRREQUISITN_STFRS.MoreResults
                SELRREQUISITN_STFRS.Requery
            End If
            NEWFLAG = False
            If SELRREQUISITN_STFRS.RowCount > 0 Then
                txt_rowcount.Text = "Total Number of Requisition for Which Sample have to Collect    " & SELRREQUISITN_STFRS.RowCount
                i = 1
                MSFlexGrid1.Rows = 1
                While Not SELRREQUISITN_STFRS.EOF
                    If SELPATHNO_FLAG = False Then
                        SELPATHNO_QR(0) = SELRREQUISITN_STFRS("REQNNO")
                        Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                        SELPATHNO_FLAG = True
                    Else
                        SELPATHNO_QR(0) = SELRREQUISITN_STFRS("REQNNO")
                        SELPATHNO_RS.MoreResults
                        SELPATHNO_RS.Requery
                    End If
                    If Trim(SELRREQUISITN_STFRS("ser_code")) <> "8888" Then
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = SELRREQUISITN_STFRS("SER_CODE")
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = SELRREQUISITN_STFRS("SER_CODE")
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_STFRS("CASENO") & vbTab & req_chglen(SELRREQUISITN_STFRS("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELRREQUISITN_STFRS("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                            i = i + 1
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_STFRS("CASENO") & vbTab & req_chglen(SELRREQUISITN_STFRS("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELRREQUISITN_STFRS("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                        i = i + 1
                    End If
                    SELRREQUISITN_STFRS.MoveNext
                Wend
                MSFlexGrid1.row = 1
                'MSFlexGrid1.col = 5
                'Commented and added by Shraddha on 21/11/2011
                MSFlexGrid1.col = 3
                chk_collected.Width = MSFlexGrid1.CellWidth
                chk_collected.Height = MSFlexGrid1.CellHeight
                chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                chk_collected.Value = 0
                chk_collected.Visible = True
            Else
                MsgBox "There are No Pending Requisition No."
             End If
        End If
       End If
End If
     'End If
'End If

If viewflag = True Then
    viewflag = FLASE
    If SELREQNOFLAG = False Then
        'SELREQNO_QR(0) = Trim(msk_caseno.FormattedText)
        SELREQNO_STFQR(0) = "P"
        SELREQNO_STFQR(1) = "C"
        SELREQNO_STFQR(2) = GET_REVDATE(msk_fromdt.FormattedText)
        SELREQNO_STFQR(3) = GET_REVDATE(msk_todt.FormattedText)
        SELREQNO_STFQR(4) = Trim(cmb_location.Text)

        Set SELREQNO_STFRS = SELREQNO_STFQR.OpenResultset(1, 3)
        SELREQNOFLAG = True
    Else
        'SELREQNO_QR(0) = msk_caseno.FormattedText
        SELREQNO_STFQR(0) = "P"
        SELREQNO_STFQR(1) = "C"
        SELREQNO_STFQR(2) = GET_REVDATE(msk_fromdt.FormattedText)
        SELREQNO_STFQR(3) = GET_REVDATE(msk_todt.FormattedText)
        SELREQNO_STFQR(4) = Trim(cmb_location.Text)
        SELREQNO_STFRS.MoreResults
        SELREQNO_STFRS.Requery
    End If
    If SELREQNO_STFRS.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELREQNO_STFRS.EOF
             If SELPATHNO_FLAG = False Then
                SELPATHNO_QR(0) = SELREQNO_STFRS("REQNNO")    '                SELRREQUISITN_RSALL("REQNNO")
                Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                SELPATHNO_FLAG = True
            Else
                SELPATHNO_QR(0) = SELREQNO_STFRS("REQNNO") ' SELRREQUISITN_RSALL("REQNNO")
                SELPATHNO_RS.MoreResults
                SELPATHNO_RS.Requery
            End If
             If SELMA101FLAG = False Then
                 SELMA101_QR(0) = SELREQNO_STFRS("SER_CODE")
                 Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                 SELMA101FLAG = True
             Else
                 SELMA101_QR(0) = SELREQNO_STFRS("SER_CODE")
                 SELMA101_RS.MoreResults
                 SELMA101_RS.Requery
            End If
            If SELMA101_RS.RowCount > 0 Then
                If Trim(SELREQNO_STFRS("REQNSTAT")) = "C" Then
                    MSFlexGrid1.AddItem i & vbTab & SELREQNO_STFRS("MODALITYCD") & vbTab & req_chglen(SELREQNO_STFRS("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELREQNO_STFRS("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                Else
                    MSFlexGrid1.AddItem i & vbTab & SELREQNO_STFRS("MODALITYCD") & vbTab & req_chglen(SELREQNO_STFRS("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELREQNO_STFRS("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                End If
            i = i + 1
            End If
            SELREQNO_STFRS.MoveNext
        Wend
        chk_collected.Visible = False
    Else
        MsgBox "There are No Pending Requisition No."
    End If
End If
End Sub
Sub ALLMODALITY()
If SELALLMODALITY_FLAGVIEW = False Then
    SELALLMODALITY_QRVIEW(0) = msk_fromdt.FormattedText
    SELALLMODALITY_QRVIEW(1) = msk_todt.FormattedText
    'SELALLMODALITY_QRVIEW(2) = "C"
    'SELALLMODALITY_QRVIEW(3) = "C"
    SELALLMODALITY_QRVIEW(2) = dept
    'i = InStr(cmb_modality.Text, "$")
    'SELRREQUISITN_QRVIEW(5) = Mid(cmb_modality, i + 1, 2)
    SELALLMODALITY_QRVIEW(3) = Trim(cmb_location.Text)
    Set SELALLMODALITY_RSVIEW = SELALLMODALITY_QRVIEW.OpenResultset(1, 3)
    SELALLMODALITY_FLAGVIEW = True
Else
    SELALLMODALITY_QRVIEW(0) = msk_fromdt.FormattedText
    SELALLMODALITY_QRVIEW(1) = msk_todt.FormattedText
    'SELALLMODALITY_QRVIEW(2) = "C"
    'SELALLMODALITY_QRVIEW(3) = "C"
    SELALLMODALITY_QRVIEW(2) = dept
    'i = InStr(cmb_modality.Text, "$")
    'SELRREQUISITN_QRVIEW(5) = Mid(cmb_modality, i + 1, 2)
    SELALLMODALITY_QRVIEW(3) = Trim(cmb_location.Text)
    SELALLMODALITY_RSVIEW.MoreResults
    SELALLMODALITY_RSVIEW.Requery
End If
If SELALLMODALITY_RSVIEW.RowCount > 0 Then
    txt_rowcount.Text = "Total Number of Requisition for Which Sample Collected    " & SELALLMODALITY_RSVIEW.RowCount
    i = 1
    MSFlexGrid1.Rows = 1
    While Not SELALLMODALITY_RSVIEW.EOF
        If SELPATHNO_FLAG = False Then
            SELPATHNO_QR(0) = SELALLMODALITY_RSVIEW("REQNNO")
            Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
            SELPATHNO_FLAG = True
        Else
            SELPATHNO_QR(0) = SELALLMODALITY_RSVIEW("REQNNO")
            SELPATHNO_RS.MoreResults
            SELPATHNO_RS.Requery
        End If
        If SELALLMODALITY_RSVIEW("SER_CODE") <> "8888" Then
            If SELMA101FLAG = False Then
                SELMA101_QR(0) = SELALLMODALITY_RSVIEW("SER_CODE")
                Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                SELMA101FLAG = True
            Else
                SELMA101_QR(0) = SELALLMODALITY_RSVIEW("SER_CODE")
                SELMA101_RS.MoreResults
                SELMA101_RS.Requery
            End If
            If SELMA101_RS.RowCount > 0 Then
                If Trim(SELALLMODALITY_RSVIEW("REQNSTAT")) = "C" Then
                    MSFlexGrid1.AddItem i & vbTab & SELALLMODALITY_RSVIEW("CASENO") & vbTab & req_chglen(SELALLMODALITY_RSVIEW("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELALLMODALITY_RSVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                Else
                    MSFlexGrid1.AddItem i & vbTab & SELALLMODALITY_RSVIEW("CASENO") & vbTab & req_chglen(SELALLMODALITY_RSVIEW("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELALLMODALITY_RSVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                End If
                i = i + 1
            End If
        Else
            If Trim(SELALLMODALITY_RSVIEW("REQNSTAT")) = "C" Then
               MSFlexGrid1.AddItem i & vbTab & SELALLMODALITY_RSVIEW("CASENO") & vbTab & req_chglen(SELALLMODALITY_RSVIEW("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELALLMODALITY_RSVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
            Else
               MSFlexGrid1.AddItem i & vbTab & SELALLMODALITY_RSVIEW("CASENO") & vbTab & req_chglen(SELALLMODALITY_RSVIEW("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELALLMODALITY_RSVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
            End If
            i = i + 1
        End If
        SELALLMODALITY_RSVIEW.MoveNext
    Wend
    chk_collected.Visible = False
Else
    MsgBox "There are No Pending Requisition No."
End If
End Sub
Private Sub NEWALLMODALITY()
If SELRREQUISITNALL_FLAG = False Then
    SELRREQUISITNALL_QR(0) = msk_fromdt.FormattedText
    SELRREQUISITNALL_QR(1) = msk_todt.FormattedText
    SELRREQUISITNALL_QR(2) = "P"
    SELRREQUISITNALL_QR(3) = "P"
    SELRREQUISITNALL_QR(4) = dept
    'i = InStr(cmb_modality.Text, "$")
    'SELRREQUISITN_QR(5) = Mid(cmb_modality, i + 1, 2)
    'SELRREQUISITN_QR(6) = Trim(cmb_location.Text)
    Set SELRREQUISITNALL_RS = SELRREQUISITNALL_QR.OpenResultset(1, 3)
Else
    SELRREQUISITNALL_QR(0) = msk_fromdt.FormattedText
    SELRREQUISITNALL_QR(1) = msk_todt.FormattedText
    SELRREQUISITNALL_QR(2) = "P"
    SELRREQUISITNALL_QR(3) = "P"
    SELRREQUISITNALL_QR(4) = dept
    'i = InStr(cmb_modality.Text, "$")
    'SELRREQUISITN_QR(5) = Mid(cmb_modality, i + 1, 2)
    'SELRREQUISITN_QR(6) = Trim(cmb_location.Text)
    SELRREQUISITNALL_RS.MoreResults
    SELRREQUISITNALL_RS.Requery
End If
NEWFLAG = False
If SELRREQUISITNALL_RS.RowCount > 0 Then
    txt_rowcount.Text = "Total Number of Requisition for Which Sample have to Collect    " & SELRREQUISITNALL_RS.RowCount
    i = 1
    MSFlexGrid1.Rows = 1
    While Not SELRREQUISITNALL_RS.EOF
        If SELPATHNO_FLAG = False Then
            SELPATHNO_QR(0) = SELRREQUISITNALL_RS("REQNNO")
            Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
            SELPATHNO_FLAG = True
        Else
            SELPATHNO_QR(0) = SELRREQUISITNALL_RS("REQNNO")
            SELPATHNO_RS.MoreResults
            SELPATHNO_RS.Requery
        End If
        If Trim(SELRREQUISITNALL_RS("ser_code")) <> "8888" Then
            If SELMA101FLAG = False Then
                SELMA101_QR(0) = SELRREQUISITNALL_RS("SER_CODE")
                Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                SELMA101FLAG = True
             Else
                SELMA101_QR(0) = SELRREQUISITNALL_RS("SER_CODE")
                SELMA101_RS.MoreResults
                SELMA101_RS.Requery
            End If
            If SELMA101_RS.RowCount > 0 Then
                MSFlexGrid1.AddItem i & vbTab & SELRREQUISITNALL_RS("CASENO") & vbTab & req_chglen(SELRREQUISITNALL_RS("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELRREQUISITNALL_RS("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                i = i + 1
            End If
        Else
            MSFlexGrid1.AddItem i & vbTab & SELRREQUISITNALL_RS("CASENO") & vbTab & req_chglen(SELRREQUISITNALL_RS("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELRREQUISITNALL_RS("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
            i = i + 1
        End If
        SELRREQUISITNALL_RS.MoveNext
    Wend
    MSFlexGrid1.row = 1
    'MSFlexGrid1.col = 5
    'Commented and added by Shraddha on 21/11/2011
    MSFlexGrid1.col = 3
    chk_collected.Width = MSFlexGrid1.CellWidth
    chk_collected.Height = MSFlexGrid1.CellHeight
    chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    chk_collected.Value = 0
    chk_collected.Visible = True
Else
    MsgBox "There are No Pending Requisition No."
End If

End Sub

Private Sub Timer1_Timer()
'COMMENTED BY CHARULATA ON 14/09/2010 FOR LESS NO. OF CALLS TO SERVER
'datetime
'txt_time.Text = time1
End Sub
Public Sub BARCODECNT()
On Error Resume Next

Dim STRIP As String

STRIP = Winsock1.LocalIP

rdb.BeginTrans
INS_BARCODE(0) = REQNO
INS_BARCODE(1) = Trim(msk_caseno.Text)
INS_BARCODE(2) = c
INS_BARCODE(3) = STRIP
INS_BARCODE(4) = DATE11
INS_BARCODE(5) = time1
INS_BARCODE(6) = Trim(cmb_location.Text)
INS_BARCODE.Execute
rdb.CommitTrans

Exit Sub

MSG:
    rdb.RollbackTrans
    'MsgBox (Err.Description)

End Sub
Public Sub NEWSTFALLMODALITY()
If SELRREQUISITNALL_STFFLAGALL = False Then
    SELRREQUISITNALL_STFQR(0) = msk_fromdt.FormattedText
    SELRREQUISITNALL_STFQR(1) = msk_todt.FormattedText
    SELRREQUISITNALL_STFQR(2) = "P"
    SELRREQUISITNALL_STFQR(3) = "P"
    SELRREQUISITNALL_STFQR(4) = dept
    'i = InStr(cmb_modality.Text, "$")
    'SELRREQUISITN_QR(5) = Mid(cmb_modality, i + 1, 2)
    'SELRREQUISITN_QR(6) = Trim(cmb_location.Text)
    Set SELRREQUISITNALL_STFRS = SELRREQUISITNALL_STFQR.OpenResultset(1, 3)
Else
    SELRREQUISITNALL_STFQR(0) = msk_fromdt.FormattedText
    SELRREQUISITNALL_STFQR(1) = msk_todt.FormattedText
    SELRREQUISITNALL_STFQR(2) = "P"
    SELRREQUISITNALL_STFQR(3) = "P"
    SELRREQUISITNALL_STFQR(4) = dept
    'i = InStr(cmb_modality.Text, "$")
    'SELRREQUISITN_QR(5) = Mid(cmb_modality, i + 1, 2)
    'SELRREQUISITN_QR(6) = Trim(cmb_location.Text)
    SELRREQUISITNALL_STFRS.MoreResults
    SELRREQUISITNALL_STFRS.Requery
End If
NEWFLAG = False
If SELRREQUISITNALL_STFRS.RowCount > 0 Then
    txt_rowcount.Text = "Total Number of Requisition for Which Sample have to Collect    " & SELRREQUISITNALL_STFRS.RowCount
    i = 1
    MSFlexGrid1.Rows = 1
    While Not SELRREQUISITNALL_STFRS.EOF
        If SELPATHNO_FLAG = False Then
            SELPATHNO_QR(0) = SELRREQUISITNALL_STFRS("REQNNO")
            Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
            SELPATHNO_FLAG = True
        Else
            SELPATHNO_QR(0) = SELRREQUISITNALL_STFRS("REQNNO")
            SELPATHNO_RS.MoreResults
            SELPATHNO_RS.Requery
        End If
        If Trim(SELRREQUISITNALL_STFRS("ser_code")) <> "8888" Then
            If SELMA101FLAG = False Then
                SELMA101_QR(0) = SELRREQUISITNALL_STFRS("SER_CODE")
                Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                SELMA101FLAG = True
             Else
                SELMA101_QR(0) = SELRREQUISITNALL_STFRS("SER_CODE")
                SELMA101_RS.MoreResults
                SELMA101_RS.Requery
            End If
            If SELMA101_RS.RowCount > 0 Then
                MSFlexGrid1.AddItem i & vbTab & SELRREQUISITNALL_STFRS("CASENO") & vbTab & req_chglen(SELRREQUISITNALL_STFRS("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELRREQUISITNALL_STFRS("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                i = i + 1
            End If
        Else
            MSFlexGrid1.AddItem i & vbTab & SELRREQUISITNALL_STFRS("CASENO") & vbTab & req_chglen(SELRREQUISITNALL_STFRS("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELRREQUISITNALL_STFRS("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
            i = i + 1
        End If
        SELRREQUISITNALL_STFRS.MoveNext
    Wend
    MSFlexGrid1.row = 1
    'MSFlexGrid1.col = 5
    'Commented and added by Shraddha on 21/11/2011
    MSFlexGrid1.col = 3
    chk_collected.Width = MSFlexGrid1.CellWidth
    chk_collected.Height = MSFlexGrid1.CellHeight
    chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    chk_collected.Value = 0
    chk_collected.Visible = True
Else
    MsgBox "There are No Pending Requisition No."
End If

End Sub

Public Sub Printlbl()            ', caseno As String, patname As String, coll_centre As String, REC_DATE As String, rec_time)
''BAR_CD1 = BAR_CD
'            BAR_CD1 = Trim(BAR_CD1)
'            bar1 = Mid(BAR_CD1, 1, 2)
'            bar2 = Mid(BAR_CD1, 3, 8)
'            bar3 = Mid(BAR_CD1, 11, 1)
'bar4 = mid(bard
'''   NEW LABEL finalised on 30/07/2004
f1 = 1

'Print #f1, "{XC00;0080,0091,A,3,02,0,0160,+0000000000,000,0,00|}"

'ORIGINAL CODE  OLD CODES
'Print #f1, "{XB00;0080,0091,A,3,02,0,0060,+0000000000,000,0,00|}"
'Print #f1, "{PC000;0140,0031,05,05,I,00,B|}"
'Print #f1, "{PC001;0030,0066,05,05,H,00,B|}"
'Print #f1, "{PC002;0089,0187,1,1,H,00,B|}" ' her
'Print #f1, "{PC003;0019,0236,05,05,H,00,B|}"
'Print #f1, "{PC004;0231,0236,05,05,H,00,B|}"
'Print #f1, "{PC005;0041,0143,05,05,I,33,B|}"
'Print #f1, "{PC006;0336,0174,05,05,H,33,B|}"
'Print #f1, "{PC007;0371,0154,05,05,I,33,B|}"
''ENDED ORIGINAL
'Print #f1, "{RB00;>6" + bar1 + ">5" + bar2 + ">6" + bar3 + "|}"
'Print #f1, "{RC000;" + CASENO + "|}"
'Print #f1, "{RC001;" + PATNAME + "|}"
'Print #f1, "{RC002;" + BAR_CD1 + " |}"
''Print #f1, "{RC003;12335678901234567890|}"
'Print #f1, "{RC003;" + coll_centre + "|}"    ' Actual Collection centre to added later when Barcode printers are installed at collection centre
'Print #f1, "{RC004;Biochemistay Lab|}"
'Print #f1, "{RC005;TMH|}"
'Print #f1, "{RC006;" + REC_DATE + "|}"
'Print #f1, "{RC007;" + REC_TIME + "|}"
'Print #f1, "{XS;I,0001,0002C4201|}"
'   OLD CODES
' NEW CODES
'TEC BARCODE PRINTER JOB START
'<xpml><page></xpml>{D0281,0380,0251|}
'<xpml></page></xpml><xpml><page></xpml>{C|}
'old barcode below upto pc0007

''commented by milind 03-07-2007

'Print #f1, "{XB00;0030,0040,A,3,03,0,0170,+0000000000,000,0,00|}"
'Print #f1, "{PC000;0281,0019,05,05,J,00,B|}"
'Print #f1, "{PC001;0020,0036,05,05,H,00,B|}"
'Print #f1, "{PC002;0131,0235,05,05,J,00,B|}"
'Print #f1, "{PC003;0320,0240,05,05,J,00,B|}"
'Print #f1, "{PC004;0030,0019,05,05,J,00,B|}"
'Print #f1, "{PC005;0091,0016,05,05,H,00,B|}"
'Print #f1, "{PC006;0190,0016,05,05,H,00,B|}"
'Print #f1, "{PC007;0020,0240,05,05,J,00,B|}"

'Print #f1, "{XB00;0080,0095,A,3,03,0,0100,+0000000000,000,0,00|}"
'Print #f1, "{PC000;0281,0035,08,08,J,00,B|}"
'Print #f1, "{PC001;0020,0080,05,06,J,00,B|}"
'Print #f1, "{PC002;0131,0230,07,07,J,00,B|}"
'Print #f1, "{PC003;0390,0235,07,07,J,00,B|}"
'Print #f1, "{PC004;0025,0027,07,07,J,00,B|}"
'Print #f1, "{PC005;0101,0016,07,07,H,00,B|}"
'Print #f1, "{PC006;0101,0045,07,07,H,00,B|}"
'Print #f1, "{PC007;0020,0235,05,05,J,00,B|}"
'Print #f1, "{PC008;0020,0235,08,08,J,00,B|}"
'Print #f1, "{RB00;>6" + bar1 + ">5" + bar2 + ">6" + bar3 + "|}"
''Print #f1, "{RB00;>6BS>504048725>6D|}"
'Print #f1, "{RC000;" + CASENO + "|}"
''Print #f1, "{RC000;TT/01500|}"
'Print #f1, "{RC001;" + PATNAME + "|}"
''Print #f1, "{RC001;Mrs. MMMMMM OIOIPOI KXKLJKLJK|}"
'Print #f1, "{RC002;" + BAR_CD1 + " |}"
''Print #f1, "{RC002;BS04048725D|}"
'Print #f1, "{RC003;FEC|}"
'Print #f1, "{RC004;TMH|}"
'Print #f1, "{RC005;" + REC_DATE + "|}"
''Print #f1, "{RC005;30H07/2004|}"
'Print #f1, "{RC006;" + REC_TIME + "|}"
''Print #f1, "{RC006;11:05|}"
'Print #f1, "{RC007;" + "" + "|}"
'Print #f1, "{RC008;" + LABNO + "|}"
'Print #f1, "{XS;I,0001,0002C4201|}"
'''upto here

'new code added by milind 03-07-07

Print #f1, "{D0205,0480,0180|}"
Print #f1, "{AY;+05,0|}"
Print #f1, "{C|}"
Print #f1, "{PC000;0010,0023,05,05,K,00,B|}"
Print #f1, "{XB00;0060,0030,9,3,03,0,0070,+0000000000,000,0,00|}"
Print #f1, "{PC001;0119,0128,15,15,G,00,B|}"
Print #f1, "{PC002;0391,0164,1,1,H,00,B|}"
Print #f1, "{PC003;0160,0158,15,15,G,00,B|}"
Print #f1, "{PC004;0020,0164,1,1,H,00,B|}"
Print #f1, "{RC000;" + PATNAME + "|}"
'Print #f1, "{RB00;" + BAR_CD1 + "|}"
'Print #f1, "{RC001;" + BAR_CD1 + "|}"
Print #f1, "{RC002;FEC|}"
Print #f1, "{RC003;" + CASENO + "|}"
''Print #f1, "{RC004;" + LABNO + "|}"
Print #f1, "{XS;I,0001,0002C2201|}"


'upto here

' TEC BARCODE PRINTER PAGE END
' <xpml></page></xpml>TEC BARCODE PRINTER JOB END
' <xpml><end/></xpml>


''' NEW CODES





'''' NEW LABEL END


End Sub

'Public Sub GETLABNO()
'
'
'
'                        If SELDMULPROCE_FLAG1 = False Then
'                            seldmulproce_qr1(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
'                            Set seldmulproce_rs1 = seldmulproce_qr1.OpenResultset(1, 3)
'                            SELDMULPROCE_FLAG1 = True
'                        Else
'                            seldmulproce_qr1(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
'                            seldmulproce_rs1.MoreResults
'                            seldmulproce_rs1.Requery
'                        End If
'
'
'
'                        REQNO1 = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
'
'                        If Not seldmulproce_rs1.EOF Then
'
'                            Do While Not seldmulproce_rs1.EOF
'
'
'                              If seldmulproce_rs1("PRO_CODE") = "F860" Or seldmulproce_rs1("PRO_CODE") = "F875" Or seldmulproce_rs1("PRO_CODE") = "F908" Or seldmulproce_rs1("PRO_CODE") = "F903" Or seldmulproce_rs1("PRO_CODE") = "F949" Or seldmulproce_rs1("PRO_CODE") = "F981" Or seldmulproce_rs1("PRO_CODE") = "F009" Or seldmulproce_rs1("PRO_CODE") = "F901" Or seldmulproce_rs1("PRO_CODE") = "F902" Or seldmulproce_rs1("PRO_CODE") = "F995" Or seldmulproce_rs1("PRO_CODE") = "F022" Then
'
'                                    If seldelprosubflag = False Then
'                                        seldelprosub_qr(0) = Trim(seldmulproce_rs1("PRO_CODE"))
'                                        Set seldelprosub_rs = seldelprosub_qr.OpenResultset(1, 3)
'                                        seldelprosubflag = True
'                                    Else
'                                        seldelprosub_qr(0) = Trim(seldmulproce_rs1("PRO_CODE"))
'                                        seldelprosub_rs.MoreResults
'                                        seldelprosub_rs.Requery
'                                    End If
'
'                                    If seldelprosub_rs.RowCount > 0 Then
'
'                                        While Not seldelprosub_rs.EOF
'
'                                        Dim REQ111t, P_CODE111t, SCODE111t
'                                        REQ111t = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
'                                        SAMPLECODE_QR1(0) = Trim(seldelprosub_rs("PROCODESUB"))
'                                        Set SAMPLECODE_RS1 = SAMPLECODE_QR1.OpenResultset(1, 3)
'                                        If SAMPLECODE_RS1.RowCount > 0 Then
'
'                                        SCODE111t = SAMPLECODE_RS1!SAMPLECD
'                                        Else
'                                        SCODE111t = " "
'                                        End If
'
'                                        REQ22 = Mid$(REQ111t, 4)
'                                        SCODE111t = REQ22 & SCODE111t
'
'                                        insdinterface_qr1(0) = MSFlexGrid1.TextMatrix(i, 2)
'                                        insdinterface_qr1(1) = REQ111t
'                                        insdinterface_qr1(2) = SCODE111t
'                                        insdinterface_qr1(3) = seldelprosub_rs("PROCODESUB")
'                                        insdinterface_qr1(4) = "1"
'                                        insdinterface_qr1(5) = time1
'                                        insdinterface_qr1(6) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
'                                        insdinterface_qr1.Execute
'
'                                        seldelprosub_rs.MoveNext
'                                    Wend
'
'                                   End If 'end if >0
'
'                                Else
'
'                                ''if not groupcode
'
'                                        Dim REQ1111t, P_CODE1111t, SCODE1111t
'                                        REQ1111t = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
'                                        SAMPLECODE_QR1(0) = Trim(seldmulproce_rs1("PRO_CODE"))
'                                        Set SAMPLECODE_RS1 = SAMPLECODE_QR1.OpenResultset(1, 3)
'                                        If SAMPLECODE_RS1.RowCount > 0 Then
'
'                                        SCODE1111t = SAMPLECODE_RS1!SAMPLECD
'                                        Else
'                                        SCODE1111t = " "
'                                        End If
'
'                                        REQ222 = Mid$(REQ1111t, 4)
'                                        SCODE1111t = REQ222 & SCODE1111t
'
'                                        insdinterface_qr1(0) = MSFlexGrid1.TextMatrix(i, 2)
'                                        insdinterface_qr1(1) = REQ1111t
'                                        insdinterface_qr1(2) = SCODE1111t
'                                        insdinterface_qr1(3) = seldmulproce_rs1("PRO_CODE")
'                                        insdinterface_qr1(4) = "1"
'                                        insdinterface_qr1(5) = time1
'                                        insdinterface_qr1(6) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
'                                        insdinterface_qr1.Execute
'
'
'                                End If
'
'                            seldmulproce_rs1.MoveNext
'                        Loop
'                        End If
'
'
'
'End Sub

Public Sub MYDELAY(inttime As String)



Dim sttime
Dim endtime1 As Long
Dim TESTTIME As Long


'sttime = Timer

endtime1 = Timer + inttime
TESTTIME = Timer

Do While Timer < endtime1
 '   Text1 = endtime1
 '   Text2 = TESTTIME
    DoEvents
  '  TESTTIME = Timer
Loop


End Sub






