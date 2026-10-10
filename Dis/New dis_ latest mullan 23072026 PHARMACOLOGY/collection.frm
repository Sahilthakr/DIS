VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form collection 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Collection Centre"
   ClientHeight    =   9345
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1440
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   1080
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   10320
      TabIndex        =   27
      Top             =   1080
      Width           =   1215
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
      Left            =   10320
      TabIndex        =   26
      Top             =   720
      Width           =   1215
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
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   7680
      Width           =   11775
      Begin VB.CommandButton Command1 
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
         Height          =   375
         Left            =   9840
         TabIndex        =   40
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CMD_REFRESH 
         Caption         =   "&Refresh"
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
         TabIndex        =   17
         Top             =   240
         Width           =   855
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
         Left            =   240
         MouseIcon       =   "collection.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   240
         Width           =   975
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
         Left            =   1440
         MouseIcon       =   "collection.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   240
         Width           =   975
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
         Left            =   2640
         MouseIcon       =   "collection.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   240
         Width           =   975
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
         Left            =   3960
         MouseIcon       =   "collection.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   240
         Width           =   855
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
         Left            =   5160
         MouseIcon       =   "collection.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   240
         Width           =   855
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
         Left            =   6240
         MouseIcon       =   "collection.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   240
         Width           =   855
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
         Left            =   7440
         MouseIcon       =   "collection.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   240
         Width           =   855
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
      TabIndex        =   24
      Top             =   3720
      Width           =   11775
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
         TabIndex        =   37
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CheckBox chk_collected 
         BackColor       =   &H8000000E&
         Caption         =   "Sample Collected"
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
         Left            =   8040
         TabIndex        =   10
         Top             =   720
         Width           =   2175
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3135
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   5530
         _Version        =   393216
         Rows            =   3
         Cols            =   10
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         ForeColorFixed  =   192
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"collection.frx":1546
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
         Left            =   1200
         TabIndex        =   36
         Top             =   3120
         Width           =   9375
      End
   End
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
      Height          =   2055
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   11775
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
         Left            =   7920
         TabIndex        =   4
         Top             =   240
         Width           =   2895
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
         Left            =   7920
         TabIndex        =   5
         Top             =   720
         Width           =   2895
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
         TabIndex        =   18
         Top             =   1560
         Width           =   3975
      End
      Begin MSMask.MaskEdBox msk_caseno 
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1575
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
            TabIndex        =   32
            Top             =   480
            Width           =   1095
         End
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
            TabIndex        =   31
            Top             =   240
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
         Left            =   5160
         TabIndex        =   2
         Top             =   240
         Width           =   1215
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
         Left            =   5160
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
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
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin MSMask.MaskEdBox msk_todt 
         Height          =   255
         Left            =   8760
         TabIndex        =   7
         Top             =   1560
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
         Left            =   6360
         TabIndex        =   6
         Top             =   1560
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
         TabIndex        =   38
         Top             =   840
         Width           =   735
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
         TabIndex        =   35
         Top             =   240
         Width           =   990
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
         Left            =   6600
         TabIndex        =   34
         Top             =   720
         Width           =   720
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
         TabIndex        =   33
         Top             =   1560
         Width           =   495
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
         TabIndex        =   30
         Top             =   240
         Width           =   1710
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
         Left            =   5160
         TabIndex        =   23
         Top             =   1560
         Width           =   885
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
         Left            =   7920
         TabIndex        =   22
         Top             =   1560
         Width           =   705
      End
   End
   Begin VB.Label Label5 
      Caption         =   "02/11/21"
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   1530
      TabIndex        =   39
      Top             =   0
      Width           =   8595
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
      Left            =   9600
      TabIndex        =   29
      Top             =   1080
      Width           =   420
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
      Left            =   9600
      TabIndex        =   28
      Top             =   720
      Width           =   420
   End
   Begin VB.Label lbl_req 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COLLECTION CENTRE"
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
      Left            =   3960
      TabIndex        =   20
      Top             =   480
      Width           =   3225
   End
End
Attribute VB_Name = "collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
Dim SELLABNO_QR As rdoQuery
Dim SELLABNO_RS As rdoResultset
Dim BARCODE_NEW As Boolean
Dim SELTIME_QR As rdoQuery
Dim SELTIME_RS As rdoResultset
Dim SELPATH_QR As rdoQuery
Dim SELPATH_RS As rdoResultset
Dim SELPROCODE_QR1 As rdoQuery
Dim SELPROCODE_RS1 As rdoResultset
Dim REQ22 As Variant

Private Sub chk_collected_Click()
If chk_collected.Value = 1 Then
    MSFlexGrid1.Text = "COLLECTED"
    
Else
    MSFlexGrid1.Text = ""
End If
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
        chk_collected = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
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
    msk_todt.Enabled = False
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
    chk_collected.Value = 0
    MSFlexGrid1.Rows = 1
btnenabl ("INI")
Me.MousePointer = 11
DIS_MAIN.show
Me.Hide
'Unload Me
Me.MousePointer = 0
End Sub

Private Sub cmd_cancel_Click()

cmb_dept.clear
'cmb_location.clear
cmb_modality.clear
msk_fromdt.Text = ""
'Msk_todt.Text = ""
msk_todt.Text = ""
txt_rowcount.Text = ""

BARCODE_NEW = False

MSFlexGrid1.Rows = 1
If OPT_ALL.Enabled = True Then
    cmb_location.Enabled = False
    msk_fromdt.Enabled = False
    msk_todt.Enabled = False
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
    msk_todt.Enabled = False
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
End Sub

Private Sub cmd_modify_Click()
btnenabl ("NEW")
chk_collected.Value = 0
chk_collected.Visible = True
End Sub

Private Sub cmd_new_Click()
On Error GoTo MSG:
NEWFLAG = True
BARCODE_NEW = True
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
    msk_todt.Enabled = False
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

If Opt_patient.Value = True Then
    msk_caseno.Mask = ""
    msk_caseno.Text = ""
    msk_caseno.Mask = "??>/#####"
ElseIf opt_staff.Value = True Then
    msk_caseno.Mask = ""
    msk_caseno.Text = ""
    msk_caseno.Mask = "#######"
End If

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
 'On Error GoTo MSG:
 If Len(Trim(cmb_location.Text)) = 0 Then
   MsgBox " Select Sample collected At....."
 Else
 rdb.BeginTrans
''''ADDED ON 14/09/2010 FOR DATETIME
datetime
txt_date.Text = DATE11
txt_time.Text = time1

'''UPTO HERE
 
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
    ''added by charulata for inserting data in tempcollct
    Call INSERTTEMPCOLLECT(CInt(i))
    '''upto here
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
MsgBox "Record is Successfully Saved."
btnenabl ("SAVE")
End If
Exit Sub
MSG:
    rdb.RollbackTrans
    MsgBox Err.Description

End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG:
btnenabl ("VIEW")
viewflag = True
msk_fromdt.Text = P_LAST_WORKING_DAY
msk_todt.Text = DATE11
msk_fromdt.Text = mydate
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
cmb_location.SetFocus
'cmb_dept.SetFocus
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Command1_Click()
''''Dim qr_cust As rdoQuery
''''STR = "INSERT INTO DIScust8(custno,name,dob,add,balance) VALUES(?,?,?,?,?)"
''''Set qr_cust = rdb.CreateQuery("qr_cust1", STR)
''''
''''qr_cust(0) = "9999"
''''qr_cust(1) = "MILIND"
''''qr_cust(2) = "07/11/11"
''''qr_cust(3) = " airoli"
''''qr_cust(4) = "1000"
''''
''''qr_cust.Execute


'Dim row As Integer
'Dim req As String
'Dim DPT As String
'Dim PRNTSTR As String
'Dim PRNTSTR1 As String
'
'For row = 1 To MSFlexGrid1.Rows
'    req = req_chglen(MSFlexGrid1.TextMatrix(row, 2))
'    req = Mid(req, 4, 10)
'    DPT = Mid(req, 1, 3)
'    PRNTSTR = "I8,1,001" & vbCrLf & "q537" & vbCrLf & "S3" & vbCrLf & "ON,D" & vbCrLf & "JF" & vbCrLf & "j204" & vbCrLf & "WN" & vbCrLf & "D10" & vbCrLf & "ZB" & vbCrLf & "Q213,37" & vbCrLf & "N" & vbCrLf
'    PRNTSTR1 = PRNTSTR & "A15,15,0,2,1,1,N,""" & Trim(txt_name) & """" & vbCrLf
'    B15 , 47, 0, 1, 3, 6, 79, N, "BC02095156A"
'    PRNTSTR1 = PRNTSTR & "A106,131,0,3,1,1,N,""" & req & """" & vbCrLf
'    A15 , 165, 0, 3, 1, 1, N, "093"
'    PRNTSTR1 = PRNTSTR1 & "A167,166,0,3,1,1,N,""" & Trim(msk_caseno) & """" & vbCrLf
'    PRNTSTR1 = PRNTSTR & "A452,164,0,4,1,1,N,""" & DPT & """" & vbCrLf
'    PRNTSTR1 = PRNTSTR1 & "P1" & vbCrLf
'Next
'
'PRNTSTR1 = PRNTSTR & "A13,17,0,1,2,2,N,""" & CTR1 & """" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "A13,55,0,1,2,2,N,""" & CTR2 & """" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "B70,89,0,1,2,6,64,N,""" & cmbAssetID.Text & """" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "A110,159,0,1,2,2,N,""" & cmbAssetID.Text & """" & vbCrLf
'PRNTSTR1 = PRNTSTR1 & "P" & cmbCopies.Text & "" & vbCrLf
'
'Open App.PATH & "\t.txt" For Output As #1
'Print #1, PRNTSTR1
'Close #1
'Shell App.PATH & "\pnt.bat", vbHidden

'''''''''barcode BY Kanchan

'Dim QTY1 As String
'Dim BAR_CD As String
'Dim EXP_DT As String
'Dim ITEM_DESC As String
'Dim TEMPDATA As String
'Dim rack As String
'Dim batchno1 As String
'Dim itemcode1 As String
'Dim i As Long, START
'Dim J As Integer
'Dim NOL As Integer
'Dim Nos As Integer
'Dim PRNTSTR As String
'Dim PRNTSTR1 As String
'Dim row As Integer
'Dim req As String
'Dim DPT As String
'Dim agesex As String
'
'
'For i = 1 To MSFlexGrid1.Rows - 1
'
'    reqno = req_chglen(MSFlexGrid1.TextMatrix(i, 2))
'    If MSFlexGrid1.TextMatrix(i, 5) = "" Then
'
'    Else
'
'    'If Mid(reqno, 4, 2) = "SE" Then
'    If Len(msk_caseno) = 6 Then
'        f1 = FreeFile()
'        Open "c:\micro.TXT" For Output As #f1
'        Dim rs_emp As rdoResultset
'        STR = "SELECT TESTCODE FROM INTERFACE.mINTERFACE WHERE REQNO=?"
'        Set SELPROCODE_QR = rdb.CreateQuery("PROCODE", STR)
'        STR = "SELECT DISTINCT LABNO FROM INTERFACE.mINTERFACE WHERE REQNO=?"
'        Set SELLABNO_QR = rdb.CreateQuery("LABNO1", STR)
'
'        'Dim row As Integer
'
'        Dim DPT1 As String
'
'        Dim agesex1 As String
'        SELLABNO_QR(0) = reqno
'        SELLABNO_QR.Execute
'        Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
'        DPT1 = Mid(reqno, 1, 3)
'        Set rs_emp = rdb.OpenResultset("select age,sex from dislib.dreqstaffd where empcode='" & Trim(msk_caseno.FormattedText) & "'", 1, 1)
'        agesex1 = rs_emp(0) & "/" & rs_emp(1)
'
'        PRNTSTR = "I8,1,001" & vbCrLf & "q591" & vbCrLf & "S3" & vbCrLf & "ON,D" & vbCrLf & "JF" & vbCrLf & "j180" & vbCrLf & "WN" & vbCrLf & "D10" & vbCrLf & "ZB" & vbCrLf & "Q184,50" & vbCrLf & "N" & vbCrLf
'        PRNTSTR1 = PRNTSTR & "A52,11,0,2,1,1,N,""" & Trim(TXT_NAME) & """" & vbCrLf
'        PRNTSTR1 = PRNTSTR1 & "B92,41,0,1,3,6,79,N,""" & SELLABNO_RS(0) & """" & vbCrLf
'        PRNTSTR1 = PRNTSTR1 & "A194,124,0,2,1,1,N,""" & SELLABNO_RS(0) & """" & vbCrLf
'        PRNTSTR1 = PRNTSTR1 & "A26,152,0,3,1,1,N,""" & agesex1 & """" & vbCrLf
'        PRNTSTR1 = PRNTSTR1 & "A185,150,0,3,1,1,N,""" & Trim(msk_caseno) & """" & vbCrLf
'        PRNTSTR1 = PRNTSTR1 & "A499,149,0,3,1,1,N,""" & DPT1 & """" & vbCrLf
'        PRNTSTR1 = PRNTSTR1 & "P1" & vbCrLf
'        Print #f1, PRNTSTR1
'        Close #f1
'        Shell "c:\micro.bat", 0
'        SELPROCODE_QR.Close
'        SELLABNO_QR.Close
'        Exit Sub
'    Else
'        f1 = FreeFile()
'        Open "c:\micro.TXT" For Output As #f1
'        Dim rs_ta101 As rdoResultset
'
'
'
'        'STR = "SELECT TESTCODE FROM INTERFACE.mINTERFACE WHERE REQNO=?"
'        'Set SELPROCODE_QR = rdb.CreateQuery("PROCODE", STR)
'
'        STR = "SELECT SAMPLECODE.SAMPLECD FROM DISLIB.SAMPLECODE,DISLIB.DMULPROCE WHERE DMULPROCE.PRO_CODE=SAMPLECODE.SERCODE AND DMULPROCE.REQNO=?"
'        Set SELLABNO_QR = rdb.CreateQuery("SELLABNO", STR)
'
'        SELLABNO_QR(0) = reqno
'        SELLABNO_QR.Execute
'
'        Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
'        Dim c As Integer
'        c = SELLABNO_RS.RowCount
'
'                    While Not SELLABNO_RS.EOF
'                        Dim REQ111t, P_CODE111t, SCODE111t
'                        REQ111t = reqno
'
'                        SCODE111t = SELLABNO_RS(0)
'
'                        REQ22 = Mid$(REQ111t, 4)
'                        SCODE111t = REQ22 & SCODE111t
'
'
'
''        'STR = "SELECT DISTINCT LABNO FROM INTERFACE.mINTERFACE WHERE REQNO=?"
''        'Set SELLABNO_QR = rdb.CreateQuery("LABNO1", STR)
''
''
''        'SELLABNO_QR(0) = reqno
''        'SELLABNO_QR.Execute
''        'Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
'
'                        DPT = Mid(reqno, 1, 3)
'                        Set rs_ta101 = rdb.OpenResultset("select birth_date,sex from patoolslib.ta101 where case_no='" & Trim(msk_caseno.FormattedText) & "'", 1, 1)
'
'                        Dim N As Integer
'                        N = DateDiff("m", rs_ta101("BIRTH_DATE"), CStr(DATE11)) / 12
'                        agesex = N & "/" & rs_ta101(1)
'
'                        PRNTSTR = "I8,1,001" & vbCrLf & "q591" & vbCrLf & "S3" & vbCrLf & "ON,D" & vbCrLf & "JF" & vbCrLf & "j180" & vbCrLf & "WN" & vbCrLf & "D10" & vbCrLf & "ZB" & vbCrLf & "Q184,50" & vbCrLf & "N" & vbCrLf
'                        PRNTSTR1 = PRNTSTR & "A52,11,0,2,1,1,N,""" & Trim(TXT_NAME) & """" & vbCrLf
'                        PRNTSTR1 = PRNTSTR1 & "B92,41,0,1,3,6,79,N,""" & SCODE111t & """" & vbCrLf
'                        PRNTSTR1 = PRNTSTR1 & "A194,124,0,2,1,1,N,""" & SCODE111t & """" & vbCrLf
'                        PRNTSTR1 = PRNTSTR1 & "A26,152,0,3,1,1,N,""" & agesex & """" & vbCrLf
'                        PRNTSTR1 = PRNTSTR1 & "A185,150,0,3,1,1,N,""" & Trim(msk_caseno) & """" & vbCrLf
'                        PRNTSTR1 = PRNTSTR1 & "A499,149,0,3,1,1,N,""" & DPT & """" & vbCrLf
'                        PRNTSTR1 = PRNTSTR1 & "P1" & vbCrLf
'                        Print #f1, PRNTSTR1
'
'                        SELLABNO_RS.MoveNext
'                    Wend
'        Close #f1
'        Shell "c:\micro.bat", 0
'        'SELPROCODE_QR.Close
'        SELLABNO_QR.Close
'        'Exit Sub
'    End If
'    End If
'    'End If
'Next i
'
'
'
'
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

Private Sub Form_Load()
On Error GoTo MSG:
NEWFLAG = False
viewflag = False
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
'str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND REQNSTAT=? AND PATNSTAT='OP' AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ')ORDER BY CASENO"

'str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND REQNSTAT=? AND PATNSTAT='OP' AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'WAB' AND DEPT_CODE <> 'CAZ')ORDER BY CASENO"
'Commented and added by Shraddha on 02/09/10 as told by Surekha m'am
str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND REQNSTAT=? AND PATNSTAT='OP' AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'WAB' AND DEPT_CODE <> 'CAZ' and AND DEPT_CODE <> 'CBZ')ORDER BY CASENO"
Set selrreqn_opqrall = rdb.CreateQuery("selrreqn_opqrall", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNDT BETWEEN ? AND ?) AND REQNSTAT=? AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ'AND DEPT_CODE <> 'WAB')ORDER BY CASENO"
Set selrreqn_opqr = rdb.CreateQuery("selrreqn_opqr", str)


'upto here


'str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND REQNSTAT= ? AND (DEPT_CODE <>'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ') AND CASENO IN(SELECT CASE_NO FROM PATOOLSLIB.TA105 WHERE DISC_DATE IS NULL AND WARD_NAME=?) ORDER BY CASENO"
str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND REQNSTAT= ? AND (DEPT_CODE <>'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'WAB' AND DEPT_CODE <> 'CAZ') AND CASENO IN(SELECT CASE_NO FROM PATOOLSLIB.TA105 WHERE DISC_DATE IS NULL AND WARD_NAME=?) ORDER BY CASENO"
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
str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNSTAT= ? OR REQNSTAT=?) AND (REQNDT BETWEEN ? AND ? )AND PATNSTAT='ST' AND (DEPT_CODE <> 'LZZ' AND DEPT_CODE <> 'WAA' AND DEPT_CODE <> 'CAZ')"
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


'UPTO HERE

str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND (REQNSTAT=? OR REQNSTAT =?) AND (REQNDT BETWEEN ? AND ? )AND (LOCATION =? OR LOCATION=?)"
Set SELREQNO_QRVIEW = rdb.CreateQuery("SELREQNO_QR1VIEW", str)


str = "select * from patoolslib.ta101 where case_no=?"
Set selta101_qr1 = rdb.CreateQuery("selta101_qr1q", str)

str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE EMPCODE=?"
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
    P_LAST_WORKING_DAY = DateAdd("d", -1, mydate)
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
   mydate = DateAdd("d", -1, mydate)
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
MSFlexGrid1.col = 5
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
msk_todt.Enabled = False
OPT_ALL.Enabled = False
OPT_CASENO.Enabled = False
Opt_patient.Enabled = False
opt_staff.Enabled = False
msk_caseno.Enabled = False
cmb_modality.Enabled = False
cmb_dept.Enabled = False
cmd_print1.Visible = False
Exit Sub
MSG:
MsgBox Err.Description
End Sub
Public Sub btnenabl(cap As String)

If cap = "NEW" Then
    cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    cmd_modify.Enabled = False
    cmd_print.Enabled = False
    cmd_view.Enabled = False
ElseIf cap = "VIEW" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = True
    'cmd_modify.Enabled = False
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_view.Enabled = False
ElseIf cap = "SAVE" Then
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = True
    cmd_view.Enabled = True
ElseIf cap = "INI" Then
    cmd_cancel.Enabled = False
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_print.Enabled = False
    cmd_view.Enabled = True
ElseIf cap = "MODIFY" Then
    cmd_save.Enabled = True
    cmd_cancel.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    cmd_modify.Enabled = False
    cmd_print.Enabled = False
    cmd_view.Enabled = False
End If


End Sub





Private Sub MSFlexGrid1_Click()
On Error GoTo MSG:
 MSFlexGrid1.col = 5
'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then

        'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
        chk_collected = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
        chk_collected.SetFocus
        chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        If Trim(MSFlexGrid1.Text) = "COLLECTED" Then
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
On Error GoTo MSG:
If Len(Trim(msk_caseno.Text)) = 0 Then Exit Sub
If Opt_patient.Value = True Then
    CASENO = padzero(5, "C", msk_caseno.Text, 3)
    msk_caseno.Text = UCase(CASENO)
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
        txt_name = Trim(selta101_rs1("tag")) & " " & Trim(selta101_rs1("first_name")) & " " & Trim(selta101_rs1("mid_name")) & " " & Trim(selta101_rs1("last_name"))
    End If
    msk_fromdt.Enabled = True
    msk_fromdt.SetFocus
ElseIf opt_staff.Value = True Then
    STAFF
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub msk_fromdt_LostFocus()
On Error GoTo MSG:
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
    Call RETRIVE_DATA
ElseIf OPT_CASENO.Value = True And opt_staff = True Then 'Or (cmb_location.Text = "STAFF") Then
    Call RETRIV_STAFF_DATE
End If
Dim i As Integer
On Error GoTo MSG:

If NEWFLAG = True Then
   If Trim(cmb_location.Text) = "GEN.OPD" Or Trim(cmb_location.Text) = "PVT.OPD" Then
        If Trim(cmb_dept.Text) = "ALL" Then
            If selrreqnflagall = False Then
                selrreqn_opqrall(0) = msk_fromdt.FormattedText
                selrreqn_opqrall(1) = msk_todt.FormattedText
                selrreqn_opqrall(2) = "P"
                Set selrreqn_oprsall = selrreqn_opqrall.OpenResultset(1, 3)
            Else
                selrreqn_opqrall(0) = msk_fromdt.FormattedText
                selrreqn_opqrall(1) = msk_todt.FormattedText
                selrreqn_opqrall(2) = "P"
                selrreqn_oprsall.MoreResults
                selrreqn_oprsall.Requery
            End If
            NEWFLAG = False
        If selrreqn_oprsall.RowCount > 0 Then
          txt_rowcount.Text = "Total Number of Requisition for Which Sample have to Collect    " & selrreqn_oprsall.RowCount
          i = 1
          MSFlexGrid1.Rows = 1
          While Not selrreqn_oprsall.EOF
              If SELPATHNO_FLAG = False Then
                  SELPATHNO_QR(0) = selrreqn_oprsall("REQNNO")
                  Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                  SELPATHNO_FLAG = True
              Else
                  SELPATHNO_QR(0) = selrreqn_oprsall("REQNNO")
                  SELPATHNO_RS.MoreResults
                  SELPATHNO_RS.Requery
              End If
              If Trim(selrreqn_oprsall("ser_code")) <> "8888" Then
                  If SELMA101FLAG = False Then
                      SELMA101_QR(0) = selrreqn_oprsall("SER_CODE")
                      Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                      SELMA101FLAG = True
                  Else
                      SELMA101_QR(0) = selrreqn_oprsall("SER_CODE")
                      SELMA101_RS.MoreResults
                      SELMA101_RS.Requery
                  End If
                  If SELMA101_RS.RowCount > 0 Then
                      MSFlexGrid1.AddItem i & vbTab & selrreqn_oprsall("CASENO") & vbTab & req_chglen(selrreqn_oprsall("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & selrreqn_oprsall("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                      i = i + 1
                  End If
              Else
                  MSFlexGrid1.AddItem i & vbTab & selrreqn_oprsall("CASENO") & vbTab & req_chglen(selrreqn_oprsall("REQNNO")) & vbTab & "M.SERVICES" & vbTab & selrreqn_oprsall("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                  i = i + 1
              End If
              selrreqn_oprsall.MoveNext
          Wend
          MSFlexGrid1.row = 1
          MSFlexGrid1.col = 5
          chk_collected.Width = MSFlexGrid1.CellWidth
          chk_collected.Height = MSFlexGrid1.CellHeight
          chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
          chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
          chk_collected.Value = 0
          chk_collected.Visible = True
      End If
    End If
End If

   If Trim(cmb_dept.Text) = "ALL" Then
      If SELRREQUISITNFLAGALL = False Then
          SELRREQUISITN_QRALL(0) = msk_fromdt.FormattedText
          SELRREQUISITN_QRALL(1) = msk_todt.FormattedText
          SELRREQUISITN_QRALL(2) = "P"
          'SELRREQUISITN_QRALL(3) = "P"
          SELRREQUISITN_QRALL(3) = Trim(cmb_location.Text)
          Set SELRREQUISITN_RSALL = SELRREQUISITN_QRALL.OpenResultset(1, 3)
      Else
          SELRREQUISITN_QRALL(0) = msk_fromdt.FormattedText
          SELRREQUISITN_QRALL(1) = msk_todt.FormattedText
          SELRREQUISITN_QRALL(2) = "P"
          'SELRREQUISITN_QRALL(3) = "P"
          SELRREQUISITN_QRALL(3) = Trim(cmb_location.Text)
          SELRREQUISITN_RSALL.MoreResults
          SELRREQUISITN_RSALL.Requery
      End If
      NEWFLAG = False

      If SELRREQUISITN_RSALL.RowCount > 0 Then
          txt_rowcount.Text = "Total Number of Requisition for Which Sample have to Collect    " & SELRREQUISITN_RSALL.RowCount
          i = 1
          MSFlexGrid1.Rows = 1
          While Not SELRREQUISITN_RSALL.EOF
              If SELPATHNO_FLAG = False Then
                  SELPATHNO_QR(0) = SELRREQUISITN_RSALL("REQNNO")
                  Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                  SELPATHNO_FLAG = True
              Else
                  SELPATHNO_QR(0) = SELRREQUISITN_RSALL("REQNNO")
                  SELPATHNO_RS.MoreResults
                  SELPATHNO_RS.Requery
              End If
              If Trim(SELRREQUISITN_RSALL("ser_code")) <> "8888" Then
                  If SELMA101FLAG = False Then
                      SELMA101_QR(0) = SELRREQUISITN_RSALL("SER_CODE")
                      Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                      SELMA101FLAG = True
                  Else
                      SELMA101_QR(0) = SELRREQUISITN_RSALL("SER_CODE")
                      SELMA101_RS.MoreResults
                      SELMA101_RS.Requery
                  End If
                  If SELMA101_RS.RowCount > 0 Then
                      MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RSALL("CASENO") & vbTab & req_chglen(SELRREQUISITN_RSALL("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELRREQUISITN_RSALL("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                      i = i + 1
                  End If
              Else
                  MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RSALL("CASENO") & vbTab & req_chglen(SELRREQUISITN_RSALL("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELRREQUISITN_RSALL("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                  i = i + 1
              End If
              SELRREQUISITN_RSALL.MoveNext
          Wend
          MSFlexGrid1.row = 1
          MSFlexGrid1.col = 5
          chk_collected.Width = MSFlexGrid1.CellWidth
          chk_collected.Height = MSFlexGrid1.CellHeight
          chk_collected.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
          chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
          chk_collected.Value = 0
          chk_collected.Visible = True
      End If
End If
If Trim(cmb_dept.Text) <> "ALL" Then
    If Trim(cmb_modality.Text) = "" Then
        NEWALLMODALITY
    Else
        If SELRREQUISITNFLAG = False Then
            SELRREQUISITN_QR(0) = msk_fromdt.FormattedText
            SELRREQUISITN_QR(1) = msk_todt.FormattedText
            SELRREQUISITN_QR(2) = "P"
            SELRREQUISITN_QR(3) = "P"
            SELRREQUISITN_QR(4) = dept
            i = InStr(cmb_modality.Text, "$")
            SELRREQUISITN_QR(5) = Mid(cmb_modality, i + 1, 2)
            'SELRREQUISITN_QR(6) = Trim(cmb_location.Text)
            Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        Else
            SELRREQUISITN_QR(0) = msk_fromdt.FormattedText
            SELRREQUISITN_QR(1) = msk_todt.FormattedText
            SELRREQUISITN_QR(2) = "P"
            SELRREQUISITN_QR(3) = "P"
            SELRREQUISITN_QR(4) = dept
            i = InStr(cmb_modality.Text, "$")
            SELRREQUISITN_QR(5) = Mid(cmb_modality, i + 1, 2)
            'SELRREQUISITN_QR(6) = Trim(cmb_location.Text)
            SELRREQUISITN_RS.MoreResults
            SELRREQUISITN_RS.Requery
        End If
        NEWFLAG = False
        If SELRREQUISITN_RS.RowCount > 0 Then
            txt_rowcount.Text = "Total Number of Requisition for Which Sample have to Collect    " & SELRREQUISITN_RS.RowCount
            i = 1
            MSFlexGrid1.Rows = 1
            While Not SELRREQUISITN_RS.EOF
                If SELPATHNO_FLAG = False Then
                    SELPATHNO_QR(0) = SELRREQUISITN_RS("REQNNO")
                    Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                    SELPATHNO_FLAG = True
                Else
                    SELPATHNO_QR(0) = SELRREQUISITN_RS("REQNNO")
                    SELPATHNO_RS.MoreResults
                    SELPATHNO_RS.Requery
                End If
                If Trim(SELRREQUISITN_RS("ser_code")) <> "8888" Then
                    If SELMA101FLAG = False Then
                        SELMA101_QR(0) = SELRREQUISITN_RS("SER_CODE")
                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                        SELMA101FLAG = True
                     Else
                        SELMA101_QR(0) = SELRREQUISITN_RS("SER_CODE")
                        SELMA101_RS.MoreResults
                        SELMA101_RS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RS("CASENO") & vbTab & req_chglen(SELRREQUISITN_RS("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELRREQUISITN_RS("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                        i = i + 1
                    End If
                Else
                    MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RS("CASENO") & vbTab & req_chglen(SELRREQUISITN_RS("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELRREQUISITN_RS("RECPTNO") & vbTab & "" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                    i = i + 1
                End If
                SELRREQUISITN_RS.MoveNext
            Wend
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 5
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
    If Trim(cmb_dept.Text) = "ALL" Then
        If SELRREQUISITNFLAGALLVIEW = False Then
            SELRREQUISITN_QRALLVIEW(0) = msk_fromdt.FormattedText
            SELRREQUISITN_QRALLVIEW(1) = msk_todt.FormattedText
            'SELRREQUISITN_QRALLVIEW(2) = "C"
            'SELRREQUISITN_QRALLVIEW(3) = "C"
           ' SELRREQUISITN_QRALL(4) = dept
           ' i = InStr(cmb_modality.Text, "$")
            'SELRREQUISITN_QRALL(5) = Mid(cmb_modality, i + 1, 2)
            SELRREQUISITN_QRALLVIEW(2) = Trim(cmb_location.Text)
            Set SELRREQUISITN_RSALLVIEW = SELRREQUISITN_QRALLVIEW.OpenResultset(1, 3)
        Else
            SELRREQUISITN_QRALLVIEW(0) = msk_fromdt.FormattedText
            SELRREQUISITN_QRALLVIEW(1) = msk_todt.FormattedText
            'SELRREQUISITN_QRALLVIEW(2) = "C"
            'SELRREQUISITN_QRALLVIEW(3) = "C"
            'SELRREQUISITN_QRALL(4) = dept
            'i = InStr(cmb_modality.Text, "$")
            'SELRREQUISITN_QRALL(5) = Mid(cmb_modality, i + 1, 2)
            SELRREQUISITN_QRALL(2) = Trim(cmb_location.Text)
            SELRREQUISITN_RSALLVIEW.MoreResults
            SELRREQUISITN_RSALLVIEW.Requery
        End If
        If SELRREQUISITN_RSALLVIEW.RowCount > 0 Then
            txt_rowcount.Text = "Total Number of Requisition for Which Sample Collected    " & SELRREQUISITN_RSALLVIEW.RowCount
            i = 1
            MSFlexGrid1.Rows = 1
            While Not SELRREQUISITN_RSALLVIEW.EOF
                If SELPATHNO_FLAG = False Then
                    SELPATHNO_QR(0) = SELRREQUISITN_RSALLVIEW("REQNNO")
                    Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                    SELPATHNO_FLAG = True
                Else
                    SELPATHNO_QR(0) = SELRREQUISITN_RSALLVIEW("REQNNO")
                    SELPATHNO_RS.MoreResults
                    SELPATHNO_RS.Requery
                End If
                If SELRREQUISITN_RSALLVIEW("SER_CODE") <> "8888" Then
                    If SELMA101FLAG = False Then
                        SELMA101_QR(0) = SELRREQUISITN_RSALLVIEW("SER_CODE")
                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                        SELMA101FLAG = True
                    Else
                        SELMA101_QR(0) = SELRREQUISITN_RSALLVIEW("SER_CODE")
                        SELMA101_RS.MoreResults
                        SELMA101_RS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQUISITN_RSALLVIEW("REQNSTAT")) = "C" Then
                            MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RSALLVIEW("CASENO") & vbTab & req_chglen(SELRREQUISITN_RSALLVIEW("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELRREQUISITN_RSALLVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                        Else
                            MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RSALLVIEW("CASENO") & vbTab & req_chglen(SELRREQUISITN_RSALLVIEW("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELRREQUISITN_RSALLVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                        End If
                        i = i + 1
                    End If
                Else
                    If Trim(SELRREQUISITN_RSALLVIEW("REQNSTAT")) = "C" Then
                        MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RSALLVIEW("CASENO") & vbTab & req_chglen(SELRREQUISITN_RSALLVIEW("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELRREQUISITN_RSALLVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                    Else
                        MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RSALLVIEW("CASENO") & vbTab & req_chglen(SELRREQUISITN_RSALLVIEW("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELRREQUISITN_RSALLVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                    End If
                    i = i + 1
                End If
                SELRREQUISITN_RSALLVIEW.MoveNext
            Wend
            chk_collected.Visible = False
        End If
   End If
   If Trim(cmb_dept.Text) <> "ALL" Then
      If Trim(cmb_modality.Text) = "" Then
          ALLMODALITY
      Else
          If SELRREQUISITNFLAGVIEW = False Then
              SELRREQUISITN_QRVIEW(0) = msk_fromdt.FormattedText
              SELRREQUISITN_QRVIEW(1) = msk_todt.FormattedText
              'SELRREQUISITN_QRVIEW(2) = "C"
              'SELRREQUISITN_QRVIEW(3) = "C"
              SELRREQUISITN_QRVIEW(2) = dept
              i = InStr(cmb_modality.Text, "$")
              SELRREQUISITN_QRVIEW(3) = Mid(cmb_modality, i + 1, 2)
              SELRREQUISITN_QRVIEW(4) = Trim(cmb_location.Text)
              Set SELRREQUISITN_RSVIEW = SELRREQUISITN_QRVIEW.OpenResultset(1, 3)
          Else
              SELRREQUISITN_QRVIEW(0) = msk_fromdt.FormattedText
              SELRREQUISITN_QRVIEW(1) = msk_todt.FormattedText
              'SELRREQUISITN_QRVIEW(2) = "C"
              'SELRREQUISITN_QRVIEW(3) = "C"
              SELRREQUISITN_QRVIEW(2) = dept
              i = InStr(cmb_modality.Text, "$")
              SELRREQUISITN_QRVIEW(3) = Mid(cmb_modality, i + 1, 2)
              SELRREQUISITN_QRVIEW(4) = Trim(cmb_location.Text)
              SELRREQUISITN_RSVIEW.MoreResults
              SELRREQUISITN_RSVIEW.Requery
          End If
          If SELRREQUISITN_RSVIEW.RowCount > 0 Then
              txt_rowcount.Text = "Total Number of Requisition for Which Sample Collected    " & SELRREQUISITN_RSVIEW.RowCount
              i = 1
              MSFlexGrid1.Rows = 1
              While Not SELRREQUISITN_RSVIEW.EOF
                  If SELPATHNO_FLAG = False Then
                      SELPATHNO_QR(0) = SELRREQUISITN_RSVIEW("REQNNO")
                      Set SELPATHNO_RS = SELPATHNO_QR.OpenResultset(1, 3)
                      SELPATHNO_FLAG = True
                  Else
                      SELPATHNO_QR(0) = SELRREQUISITN_RSVIEW("REQNNO")
                      SELPATHNO_RS.MoreResults
                      SELPATHNO_RS.Requery
                  End If
                  If SELRREQUISITN_RSVIEW("SER_CODE") <> "8888" Then
                      If SELMA101FLAG = False Then
                          SELMA101_QR(0) = SELRREQUISITN_RSVIEW("SER_CODE")
                          Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                          SELMA101FLAG = True
                      Else
                          SELMA101_QR(0) = SELRREQUISITN_RSVIEW("SER_CODE")
                          SELMA101_RS.MoreResults
                          SELMA101_RS.Requery
                      End If
                      If SELMA101_RS.RowCount > 0 Then
                          If Trim(SELRREQUISITN_RSVIEW("REQNSTAT")) = "C" Then
                              MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RSVIEW("CASENO") & vbTab & req_chglen(SELRREQUISITN_RSVIEW("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELRREQUISITN_RSVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                          Else
                              MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RSVIEW("CASENO") & vbTab & req_chglen(SELRREQUISITN_RSVIEW("REQNNO")) & vbTab & Trim(SELMA101_RS("SER_DET")) & vbTab & SELRREQUISITN_RSVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                          End If
                          i = i + 1
                      End If
                  Else
                      If Trim(SELRREQUISITN_RSVIEW("REQNSTAT")) = "C" Then
                         MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RSVIEW("CASENO") & vbTab & req_chglen(SELRREQUISITN_RSVIEW("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELRREQUISITN_RSVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                      Else
                         MSFlexGrid1.AddItem i & vbTab & SELRREQUISITN_RSVIEW("CASENO") & vbTab & req_chglen(SELRREQUISITN_RSVIEW("REQNNO")) & vbTab & "M.SERVICES" & vbTab & SELRREQUISITN_RSVIEW("RECPTNO") & vbTab & "COLLECTED" & vbTab & SELPATHNO_RS("PATHNO") & vbTab & Trim(SELPATHNO_RS("REFTODOC"))
                      End If
                      i = i + 1
                  End If
                  SELRREQUISITN_RSVIEW.MoveNext
              Wend
              chk_collected.Visible = False
          Else
              MsgBox "There are No Pending Requisition No."
          End If
      End If
    End If
    'viewflag = False
End If
Exit Sub
MSG:
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
    msk_caseno.Mask = "??>/#####"
    OPT_CASENO.Enabled = True
    'Opt_all.Enabled = False
    msk_fromdt.Enabled = False

    msk_todt.Enabled = False
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
    msk_todt.Enabled = False
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
        SELDREQSTAFFD_RS.MoveLast
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
    MSFlexGrid1.col = 5
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
                Set selrreqn_oprs = selrreqn_opqr.OpenResultset(1, 3)
            Else
                selrreqn_opqr(0) = msk_caseno.FormattedText
                selrreqn_opqr(1) = msk_fromdt.FormattedText
                selrreqn_opqr(2) = msk_todt.FormattedText
                selrreqn_opqr(3) = "P"
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
          MSFlexGrid1.col = 5
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
                MSFlexGrid1.col = 5
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
    SELREQNO_QRVIEW(5) = Trim(cmb_location.Text)
    SELREQNO_QRVIEW(6) = Trim(cmb_location.Text)
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
                    MSFlexGrid1.col = 5
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
                        MSFlexGrid1.col = 5
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
                MSFlexGrid1.col = 5
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
    MSFlexGrid1.col = 5
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
    MSFlexGrid1.col = 5
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
Print #f1, "{XS;I,0001,0002C4201|}"


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

Public Sub INSERTTEMPCOLLECT(rownum As Integer)
    Dim SEL_SRNO_QR As rdoQuery
    Dim SEL_MAXSRNO_QR As rdoQuery
    Dim INS_TEMPCOLLCT As rdoQuery
    Dim SEL_TA101_QR As rdoQuery
    Dim selma104_qr As rdoQuery
    Dim seldmulproce_qr As rdoQuery
    Dim SEL_TEMPCOLLCTR_QR As rdoQuery
    Dim upd_TEMPCOLLCTR_qr As rdoQuery
    
    Dim SEL_SRNO_RS As rdoResultset
    Dim SEL_MAXSRNO_RS As rdoResultset
    Dim SEL_TA101_RS As rdoResultset
    Dim selma104_rs As rdoResultset
    Dim seldmulproce_rs As rdoResultset
    Dim SEL_TEMPCOLLCTR_rs As rdoResultset
    Dim upd_TEMPCOLLCTR_rs As rdoResultset
    
    
    Dim name As String
    Dim age As String
    Dim SEX As String
    Dim pat_catg As String
    Dim patstat As String
    
    str = "SELECT * FROM Dislib.TEMPCOLLCT WHERE caseno=? AND SENTDATE=? and LOCATION=?"
    Set SEL_SRNO_QR = rdb.CreateQuery("", str)
    
    str = "SELECT max(SRNO) FROM Dislib.TEMPCOLLCT WHERE LOCATION=? AND COLLDATE=?"
    Set SEL_MAXSRNO_QR = rdb.CreateQuery("", str)
    
    str = "INSERT INTO DISLIB.TEMPCOLLCT (SRNO,REQNO,MODALITYCD,CASENO,NAME,SEX,AGE,IPADD,PATNSTAT,PAT_CATG,LOCATION,COLLDATE,COLLTIME,SENTDATE,SENTTIME,SER_CODE,SER_DET,SENTFLAG,CFLAG,USERID) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
    Set INS_TEMPCOLLCT = rdb.CreateQuery("", str)

    str = "select * from PATOOLSLIB.ta101 where case_no=?"
    Set SEL_TA101_QR = rdb.CreateQuery("", str)

    str = "select * from PATOOLSLIB.ma104 where case_no=?"
    Set selma104_qr = rdb.CreateQuery("", str)
                            
    str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
    Set seldmulproce_qr = rdb.CreateQuery("", str)

    str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
    Set SELMA101_QR = rdb.CreateQuery("", str)
    
    str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
    Set seldmulproce_qr = rdb.CreateQuery("", str)
    
    str = "SELECT * FROM Dislib.TEMPCOLLCT WHERE REQNO=? AND SER_CODE=?"
    Set SEL_TEMPCOLLCTR_QR = rdb.CreateQuery("", str)

    '***** Shraddha 34/09/2011 *****
    If Len(Trim(MSFlexGrid1.TextMatrix(rownum, 1))) = 6 Then
        str = "SELECT * FROM Dislib.DREQSTAFFD WHERE REQNO=?" 'DREQSTAFFD
        Set SEL_TA101_QR = rdb.CreateQuery("", str)
        SEL_TA101_QR(0) = req_chglen(Trim(MSFlexGrid1.TextMatrix(rownum, 2)))
    Else
    '*******************************
        str = "select * from PATOOLSLIB.ta101 where case_no=?"
        Set SEL_TA101_QR = rdb.CreateQuery("", str)
        SEL_TA101_QR(0) = Trim(MSFlexGrid1.TextMatrix(rownum, 1))
    End If
    Set SEL_TA101_RS = SEL_TA101_QR.OpenResultset(2, 3)
    
    str = "update Dislib.TEMPCOLLCT set CFLAG=? WHERE REQNO=? AND SER_CODE=?"
    Set upd_TEMPCOLLCTR_qr = rdb.CreateQuery("", str)
                            
    seldmulproce_qr(0) = req_chglen(Trim(MSFlexGrid1.TextMatrix(rownum, 2)))
    Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 3)
                              
      While Not seldmulproce_rs.EOF
        'name = Trim(SEL_TA101_RS("TAG")) & " " & Trim(SEL_TA101_RS("FIRST_NAME")) & " " & Trim(SEL_TA101_RS("MID_NAME")) & " " & Trim(SEL_TA101_RS("LAST_NAME"))
        If Len(Trim(MSFlexGrid1.TextMatrix(rownum, 1))) = 6 Then
            name = Trim(SEL_TA101_RS("PNAME"))
            age = Trim(SEL_TA101_RS("AGE"))
            SEX = Trim(SEL_TA101_RS("SEX"))
            pat_catg = ""
        Else
        name = Trim(SEL_TA101_RS("TAG")) & " " & Trim(SEL_TA101_RS("FIRST_NAME")) & " " & Trim(SEL_TA101_RS("LAST_NAME"))
        age = DateDiff("YYYY", SEL_TA101_RS("BIRTH_DATE"), CStr(DATE11))
        SEX = Trim(SEL_TA101_RS("SEX"))
        pat_catg = Trim(SEL_TA101_RS("PAT_CATG"))
        End If
        
        selma104_qr(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
        Set selma104_rs = selma104_qr.OpenResultset(1, 3)
          
        If selma104_rs.RowCount > 0 Then
            patstat = "IP"
        Else
            patstat = "OP"
        End If
    
    
    
        SEL_SRNO_QR(0) = Trim(MSFlexGrid1.TextMatrix(rownum, 1))
        SEL_SRNO_QR(1) = Format(DATE11, "YYYY-MM-DD")
        SEL_SRNO_QR(2) = Trim(cmb_location.Text)
        Set SEL_SRNO_RS = SEL_SRNO_QR.OpenResultset(1, 3)
        
        If Not SEL_SRNO_RS.EOF Then
            INS_TEMPCOLLCT(0) = Trim(SEL_SRNO_RS("SRNO"))
        Else
        'If first time then assign 1
            SEL_MAXSRNO_QR(0) = Trim(cmb_location.Text)
            SEL_MAXSRNO_QR(1) = Format(DATE11, "YYYY-MM-DD")
            Set SEL_MAXSRNO_RS = SEL_MAXSRNO_QR.OpenResultset(1, 3)
            
            If IsNull(SEL_MAXSRNO_RS(0)) = True Then
                INS_TEMPCOLLCT(0) = 1
            Else
                INS_TEMPCOLLCT(0) = SEL_MAXSRNO_RS(0) + 1
            End If
                
            
        End If
        '**********
        
        INS_TEMPCOLLCT(1) = req_chglen(Trim(MSFlexGrid1.TextMatrix(rownum, 2))) 'REQNO
        INS_TEMPCOLLCT(2) = Mid(Trim(MSFlexGrid1.TextMatrix(rownum, 2)), 5, 2)
        INS_TEMPCOLLCT(3) = Trim(MSFlexGrid1.TextMatrix(rownum, 1)) 'CASENO
        INS_TEMPCOLLCT(4) = name
        INS_TEMPCOLLCT(5) = SEX
        INS_TEMPCOLLCT(6) = age
        INS_TEMPCOLLCT(7) = Winsock1.LocalIP
        INS_TEMPCOLLCT(8) = patstat
        INS_TEMPCOLLCT(9) = pat_catg
        INS_TEMPCOLLCT(10) = Trim(cmb_location.Text)
        INS_TEMPCOLLCT(11) = DATE11 'COLLDATE
        INS_TEMPCOLLCT(12) = time1 'COLLTIME
        INS_TEMPCOLLCT(13) = DATE11
        INS_TEMPCOLLCT(14) = time1
       
        
        'SER_CODE,SER_DET,SENTFLAG,CFLAG
        INS_TEMPCOLLCT(15) = Trim(seldmulproce_rs("PRO_CODE"))
        '***** 13/09/2011 *****
        SELMA101_QR(0) = Trim(seldmulproce_rs("PRO_CODE"))
        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
        '**********************
        INS_TEMPCOLLCT(16) = ""
        If Not SELMA101_RS.EOF Then
            INS_TEMPCOLLCT(16) = Trim(SELMA101_RS("SER_DET"))
        End If
        INS_TEMPCOLLCT(17) = "" 'SENTFLAG
        INS_TEMPCOLLCT(18) = "Y" 'Trim(MSFlexGrid1.TextMatrix(rownum, 7)) 'CFLAG
        INS_TEMPCOLLCT(19) = Right(Trim(UID), 6) 'USERID
        
        
        SEL_TEMPCOLLCTR_QR(0) = Trim(MSFlexGrid1.TextMatrix(rownum, 1))
        SEL_TEMPCOLLCTR_QR(1) = Trim(seldmulproce_rs("PRO_CODE"))
        Set SEL_TEMPCOLLCTR_rs = SEL_TEMPCOLLCTR_QR.OpenResultset(1, 3)
        
        If SEL_TEMPCOLLCTR_rs.EOF Then
            INS_TEMPCOLLCT.Execute
        Else
        
            'Check cflag and keep it as it is.
            upd_TEMPCOLLCTR_qr(0) = "Y" 'Trim(MSFlexGrid1.TextMatrix(rownum, 7))
            upd_TEMPCOLLCTR_qr(1) = Trim(MSFlexGrid1.TextMatrix(rownum, 1))
            upd_TEMPCOLLCTR_qr(2) = Trim(seldmulproce_rs("PRO_CODE"))
            upd_TEMPCOLLCTR_qr.Execute
        End If
        seldmulproce_rs.MoveNext
    Wend
    
End Sub
