VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_ICU_micro 
   BackColor       =   &H80000003&
   Caption         =   "Pathogen Profile"
   ClientHeight    =   10905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15660
   LinkTopic       =   "Form2"
   ScaleHeight     =   10905
   ScaleWidth      =   15660
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1665
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   900
      Width           =   240
   End
   Begin VB.CommandButton cmd_icuexp 
      Caption         =   "ALL Antibiotic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11250
      TabIndex        =   40
      Top             =   5715
      Width           =   1590
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   240
      Left            =   360
      TabIndex        =   39
      Top             =   5625
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   423
      _Version        =   393216
      Cols            =   240
   End
   Begin VB.OptionButton opt_OP 
      BackColor       =   &H80000003&
      Caption         =   "OP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   38
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmd_ExpEXCEL 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13200
      TabIndex        =   37
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Frame frame_SA 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1680
      TabIndex        =   21
      Top             =   2760
      Width           =   13215
      Begin VB.ComboBox cmb_organism 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10200
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   2880
      End
      Begin VB.OptionButton Opt_organism 
         BackColor       =   &H80000003&
         Caption         =   "Organism :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8880
         TabIndex        =   35
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton Opt_antibiotic 
         BackColor       =   &H80000003&
         Caption         =   "Antibiotic :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   25
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton Opt_specimen 
         BackColor       =   &H80000003&
         Caption         =   "Specimen :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   1455
      End
      Begin VB.ComboBox cmb_antibiotic 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5880
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ComboBox cmb_specimen 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Shape Shape_SA 
         Height          =   615
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   13215
      End
   End
   Begin VB.CommandButton cmd_clear 
      BackColor       =   &H80000013&
      Caption         =   "&CLEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_back 
      BackColor       =   &H80000013&
      Caption         =   "&BACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_exit 
      BackColor       =   &H80000013&
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_print 
      BackColor       =   &H80000013&
      Caption         =   "&PRINT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_view 
      BackColor       =   &H80000013&
      Caption         =   "&VIEW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000003&
      Caption         =   "Ward"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1800
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000003&
      Caption         =   "ICU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2280
      Width           =   855
   End
   Begin VB.ComboBox cmb_ward 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      TabIndex        =   9
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton cmd_EMR 
      BackColor       =   &H0080C0FF&
      Caption         =   "&EMR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000003&
      Caption         =   "Patient"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame Frame_colour 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Caption         =   "Colour Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   13440
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
      Begin VB.Label Label8 
         BackColor       =   &H0080FF80&
         Caption         =   "Sensitive"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H008080FF&
         Caption         =   "Resistant"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080FFFF&
         Caption         =   "Intermediate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H80000003&
      Caption         =   "Unit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.ComboBox cmb_unit 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   1
      Top             =   3720
      Width           =   3135
   End
   Begin VB.ComboBox cmb_doctor 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8040
      TabIndex        =   0
      Top             =   3720
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4455
      Left            =   120
      TabIndex        =   13
      Top             =   6120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   7858
      _Version        =   393216
      Cols            =   42
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   -2147483629
      ForeColorFixed  =   0
      AllowUserResizing=   1
      FormatString    =   $"ICU_micro.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483629
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   375
      Left            =   7560
      TabIndex        =   19
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483629
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin Crystal.CrystalReport CR 
      Left            =   120
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   375
      Left            =   3480
      TabIndex        =   20
      Top             =   3120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483629
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
   Begin VB.Line Line1 
      X1              =   120
      X2              =   13320
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000003&
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   34
      Top             =   120
      Width           =   9615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000003&
      Caption         =   "To Date : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   33
      Top             =   1560
      Width           =   855
   End
   Begin VB.Shape z 
      BackColor       =   &H80000013&
      Height          =   3975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   13215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "From Date : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   32
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000003&
      Caption         =   "PATHOGEN PROFILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   31
      Top             =   720
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000003&
      Height          =   2415
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lbl_caseno 
      BackColor       =   &H80000003&
      Caption         =   "Case No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   30
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lbl_wards 
      BackColor       =   &H80000003&
      Caption         =   " Ward : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   29
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lbl_caption 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   28
      Top             =   5520
      Width           =   11535
   End
   Begin VB.Shape Shape_WSA 
      Height          =   1335
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   11655
   End
   Begin VB.Shape Shape_unit 
      Height          =   615
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   9375
   End
   Begin VB.Label lbl_unit 
      BackColor       =   &H80000003&
      Caption         =   "Unit : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   27
      Top             =   3840
      Width           =   495
   End
   Begin VB.Shape Shape_patient 
      Height          =   615
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label lbl_doctor 
      BackColor       =   &H80000003&
      Caption         =   "Doctor : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   26
      Top             =   3840
      Width           =   735
   End
End
Attribute VB_Name = "frm_ICU_micro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************
'This form has been developed in its entirety by Ms. Shraddha Pandit (Commense Date : 06/05/2009)
'*********************************************************************************
Dim var_date_f As String
Dim var_date_t As String

Dim sql3 As String

Dim sel_dreq_qr As rdoQuery
Dim sel_dreq_rs As rdoResultset

Dim UI As String
Dim sel_ward_qr As rdoQuery
Dim sel_ward_rs As rdoResultset
Dim wardname, CASENO As String
Dim i, j, n As Integer
Dim sal, FNAME, mname, lname, pname As String

Dim sql_pathogen_ward As String
Dim sql_pathogen_icu As String
Dim sql_pathogen_patient As String
Dim sql_pathogen_unit As String

Dim sel_pathogen_ward_qr  As rdoQuery
Dim sel_pathogen_ward_rs As rdoResultset

'***** Shraddha 22/11/10 *****
Dim sel_pathogen_Allward_qr  As rdoQuery
Dim sel_pathogen_Allward_rs As rdoResultset
'*****************************

Dim patho_grown As String

Dim sel_pathogen_icu_qr As rdoQuery
Dim sel_pathogen_icu_rs As rdoResultset

Dim sel_pathogen_patient_qr As rdoQuery
Dim sel_pathogen_patient_rs As rdoResultset
Dim c_no As String

Dim sel_pathogen_unit_qr As rdoQuery
Dim sel_pathogen_unit_rs As rdoResultset
Dim unit, doctor As String



Dim sel_labno_qr As rdoQuery
Dim sel_labno_rs As rdoResultset
Dim lab_no As String

Dim cno, rno, nm, or1, or2, or3, antbt, gender, nmaterial, antbt1, antbt2, antbt3 As String
Dim age, YEAR1 As Integer

Dim specimen, antibiotic, organism, Title As String
Dim title_chk As Boolean
Dim colldt As Date

Private Sub cmb_ward_LostFocus()
If Trim(cmb_ward.Text) = "ALL" Then
    frame_SA.Visible = False
Else
    frame_SA.Visible = True
End If
End Sub

Private Sub cmd_ExpEXCEL_Click()

'Added on 12/11/10 by Shraddha
     
Dim xlObject    As Excel.Application

Dim xlWB        As Excel.Workbook

Set xlObject = New Excel.Application

'This Adds a new woorkbook, you could open the workbook from file also

Set xlWB = xlObject.Workbooks.add

Clipboard.clear 'Clear the Clipboard

With MSFlexGrid1

    'Select Full Contents (You could also select partial content)
    
    .col = 0               'From first column
    
    .row = 0               'From first Row (header)
    
    .ColSel = .Cols - 1    'Select all columns
    
    .RowSel = .Rows - 1    'Select all rows
    
    Clipboard.SetText .Clip 'Send to Clipboard

End With



With xlObject.ActiveWorkbook.ActiveSheet

    .range("A1").Select 'Select Cell A1 (will paste from here, to different cells)

    .Paste              'Paste clipboard contents

End With



' This makes Excel visible

xlObject.Visible = True
  
   
End Sub

Private Sub cmd_icuexp_Click()
Dim xlObject    As Excel.Application

Dim xlWB        As Excel.Workbook

Set xlObject = New Excel.Application

Dim rs_dmicbantib As rdoResultset

'This Adds a new woorkbook, you could open the workbook from file also

Set xlWB = xlObject.Workbooks.add

If Option1.Value = True Or Option2.Value = True Then
    
    
    InsDataPathogen
    
Else
    MsgBox ("select WARD or ICU")
    Exit Sub
End If



With xlObject.ActiveWorkbook.ActiveSheet
    
    Set rs_dmicbantib = rdb.OpenResultset("SELECT antibdesc FROM DISLIB.dmicbantib where modality='BACTERIOLOGY'", 1, 1)
    
    Clipboard.clear
    Clipboard.SetText lbl_caption.Caption
    
    .range("A1:HZ1").Select
    .Paste
    
    str = "Reqno" & vbTab & "Caseno" & vbTab & "CollectionDT" & vbTab & "Specimen" & vbTab & "Org1" & vbTab
    Do While Not rs_dmicbantib.EOF
        str = str & Trim(rs_dmicbantib(0)) & "1" & vbTab
        rs_dmicbantib.MoveNext
            
    Loop
    str = str & vbTab
    str = str & "Org2" & vbTab
    rs_dmicbantib.MoveFirst
    
    Do While Not rs_dmicbantib.EOF
        str = str & Trim(rs_dmicbantib(0)) & "2" & vbTab
        rs_dmicbantib.MoveNext
            
    Loop
    str = str & vbTab & vbTab
    str = str & "Org3" & vbTab
    rs_dmicbantib.MoveFirst
    
    Do While Not rs_dmicbantib.EOF
        str = str & Trim(rs_dmicbantib(0)) & "3" & vbTab
        rs_dmicbantib.MoveNext
            
    Loop
    
    
    Clipboard.clear
    
    Clipboard.SetText str
    
    .range("A2:HZ2").Select
    
    .Paste
    
    Clipboard.clear 'Clear the Clipboard

With MSFlexGrid1

    'Select Full Contents (You could also select partial content)
    
    .col = 0               'From first column
    
    .row = 1               'From first Row (header)
    
    .ColSel = .Cols - 1    'Select all columns
    
    .RowSel = .Rows - 1    'Select all rows
    
    Clipboard.SetText .Clip 'Send to Clipboard

End With
    .range("A3").Select 'Select Cell A1 (will paste from here, to different cells)

    .Paste              'Paste clipboard contents
    
    

End With



' This makes Excel visible

xlObject.Visible = True
  
End Sub

'Dim sql_pathogen_icuANTI As String

Private Sub Form_Load()
patho_grown = ""
sql_ward = "select * from DISLIB.wards order by WARD_NAME"
Set sel_ward_qr = rdb.CreateQuery("", sql_ward)

'******* Only according to ward **********
'sql_pathogen_ward = "select * from RISLIB.pathogen where WARD_NAME=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
'*****************************************
''********** According to ward and specimen ********
'sql_pathogen_ward = "select * from RISLIB.pathogen where WARD_NAME=? AND SPECIMEN=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
'Set sel_pathogen_ward_qr = rdb.CreateQuery("", sql_pathogen_ward)
''**************************************************
'
''********** According to ward and Antibiotic 15/06/09 ********
'sql_pathogen_ward = "select * from RISLIB.pathogen where WARD_NAME=? AND ANTIBIOTIC=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
'Set sel_pathogen_ward_qr = rdb.CreateQuery("", sql_pathogen_ward)
''*************************************************************

'******* Only according to ICU ***********
'sql_pathogen_icu = "select * from RISLIB.pathogen where ICDS2='YES' AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
'*****************************************

''******* According to ICU and Specimen ***********
'sql_pathogen_icu = "select * from RISLIB.pathogen where ICDS2='YES' AND SPECIMEN=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
'Set sel_pathogen_icu_qr = rdb.CreateQuery("", sql_pathogen_icu)
''*************************************************
'
''******* According to ICU and Antibiotic 15/06/09 ***********
'sql_pathogen_icu = "select * from RISLIB.pathogen where ICDS2='YES' AND ANTIBIOTIC=? and (ANTIBIO1='S' or ANTIBIO2='S' or ANTIBIO3='S') AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
'Set sel_pathogen_icu_qr = rdb.CreateQuery("", sql_pathogen_icu)
''************************************************************
title_chk = False

'********** Only according to patient ************
sql_pathogen_patient = "select * from RISLIB.pathogen_new where caseno=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
Set sel_pathogen_patient_qr = rdb.CreateQuery("", sql_pathogen_patient)
'*************************************************

sql3 = "Select * from dislib.drequisitn where reqno=?"
Set sel_labno_qr = rdb.CreateQuery("", sql3)


'********************** Antibiotic ********************************
str = "select * from dislib.dmicbantib WHERE MODALITY='BACTERIOLOGY' AND SUBMODALI='BACTERIOLOGY' ORDER BY ANTIBDESC"
Set sel_antiqr = rdb.CreateQuery("", str)

Set sel_antirs = sel_antiqr.OpenResultset(1, 3)
cmb_antibiotic.AddItem "ALL" '03/08/09
While Not sel_antirs.EOF
    cmb_antibiotic.AddItem sel_antirs("antibdesc")
    sel_antirs.MoveNext
Wend

'******************************************************************
datetime

'***** Shraddha 08/12/09 *****
Dim org_QR As rdoQuery
Dim org_RS As rdoResultset
Dim sql_org As String

sql_org = "SELECT * FROM DISLIB.organism order by ORGANISM"
Set org_QR = rdb.CreateQuery("", sql_org)
Set org_RS = org_QR.OpenResultset(1, 3)

If Not org_RS.EOF Then

    cmb_organism.AddItem "ALL" '09/12/09

    While Not org_RS.EOF
        cmb_organism.AddItem Trim(org_RS("ORGANISM"))
        org_RS.MoveNext
    Wend
End If
'*****************************

Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False

cmb_ward.Visible = False
lbl_wards.Visible = False
MaskEdBox3.Visible = False
lbl_caseno.Visible = False
MSFlexGrid1.Visible = False
Frame_colour.Visible = False
lbl_caption.Visible = False
'cmb_specimen.Visible = False
'cmb_antibiotic.Visible = False
'lbl_specimen.Visible = False
frame_SA.Visible = False
Shape_SA.Visible = False
Shape_WSA.Visible = False
cmb_unit.Visible = False
Shape_unit.Visible = False
Shape_patient.Visible = False
lbl_unit.Visible = False
lbl_doctor.Visible = False
cmb_doctor.Visible = False

cmb_organism.Visible = False

Set sel_ward_rs = sel_ward_qr.OpenResultset(1, 3)

'***** Shraddha 22/11/10 *****
cmb_ward.AddItem "ALL"
'*****************************

While Not sel_ward_rs.EOF
               
    If Trim(sel_ward_rs(0)) <> "" Then
        cmb_ward.AddItem Trim(sel_ward_rs(0))
    End If
    sel_ward_rs.MoveNext
                
Wend
sel_ward_rs.Close

MaskEdBox1.Text = Format(Now, "dd/mm/yyyy")
MaskEdBox2.Text = Format(Now, "dd/mm/yyyy")


'************** Specimen *****************

Dim sample_QR As rdoQuery
Dim sample_RS As rdoResultset
Dim sql As String

sql = "SELECT * FROM DISLIB.microsampl order by SAMPLENAME"
Set sample_QR = rdb.CreateQuery("", sql)
Set sample_RS = sample_QR.OpenResultset(1, 3)

If Not sample_RS.EOF Then
      
    cmb_specimen.AddItem "ALL" '03/08/09
    While Not sample_RS.EOF
        
        cmb_specimen.AddItem sample_RS("SAMPLENAME")
        sample_RS.MoveNext
        
    Wend
    
End If
sample_RS.Close

'*************************** Unit ***************************
'SELECT * FROM PATOOLSLIB/MA1031 163
'UNIT_NAME  57 D=105
Dim unit_QR As rdoQuery
Dim unit_RS As rdoResultset
Dim sql_unit As String

sql_unit = "SELECT distinct(UNIT_NAME) FROM PATOOLSLIB.MA1031 order by UNIT_NAME"
Set unit_QR = rdb.CreateQuery("", sql_unit)
Set unit_RS = unit_QR.OpenResultset(1, 3)

If Not unit_RS.EOF Then
    cmb_unit.AddItem "ALL"
    While Not unit_RS.EOF
        If Trim(unit_RS(0)) <> "" Then
            cmb_unit.AddItem unit_RS(0)
        End If
            unit_RS.MoveNext
    Wend
        
End If
unit_RS.Close
'***********************************************************************

End Sub
Private Sub CMD_EMR_Click()
    Dim VAR1
    VAR1 = MaskEdBox3.FormattedText
    VAR = Shell("C:\EMR1.exe " & VAR1, vbNormalFocus)
    AppActivate VAR
End Sub
Private Sub cmd_back_Click()
'Unload Me
patho_grown = ""
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_clear_Click()
patho_grown = ""

MaskEdBox1.Text = "__/__/____"
MaskEdBox2.Text = "__/__/____"
MSFlexGrid1.clear

MaskEdBox3.Text = ""
MaskEdBox3.Visible = False
lbl_caseno.Visible = False

cmb_ward.Visible = False
lbl_wards.Visible = False
'cmb_specimen.Visible = False
'lbl_specimen.Visible = False
'cmb_antibiotic.Visible = False
frame_SA.Visible = False
Shape_SA.Visible = False
Shape_WSA.Visible = False

Frame_colour.Visible = False
lbl_caption.Visible = False
MSFlexGrid1.Visible = False
Option1.Value = False
Option2.Value = False
Option3.Value = False
Opt_specimen.Value = False
Opt_organism.Value = False

Opt_antibiotic.Value = False
cmb_ward.Text = ""
cmb_specimen.Text = ""
cmb_organism.Text = "" '09/12/09
cmb_antibiotic.Text = ""
'MaskEdBox1.SetFocus
End Sub

Private Sub cmd_exit_Click()
patho_grown = ""
End
End Sub

Private Sub MaskEdBox1_GotFocus()
'MaskEdBox1.Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub MaskEdBox2_GotFocus()
'MaskEdBox2.Text = Format(Now, "dd/mm/yyyy")
End Sub


Private Sub cmd_print_Click()
If MaskEdBox1.Text = "__/__/____" Then
    MsgBox "Please enter From date "
    MaskEdBox1.Text = "__/__/____"
    MaskEdBox1.SetFocus
    Exit Sub
ElseIf IsDate(MaskEdBox1.Text) = False Then
    MsgBox "Please enter valid From date "
    MaskEdBox1.Text = "__/__/____"
    MaskEdBox1.SetFocus
    Exit Sub
End If

If MaskEdBox2.Text = "__/__/____" Then
    MsgBox "Please enter To date "
    MaskEdBox2.Text = "__/__/____"
    MaskEdBox2.SetFocus
    Exit Sub
ElseIf IsDate(MaskEdBox2.Text) = False Then
    MsgBox "Please enter valid To date "
    MaskEdBox2.Text = "__/__/____"
    MaskEdBox2.SetFocus
    Exit Sub
End If

Dim day_cnt
day_cnt = DateDiff("d", MaskEdBox1.Text, MaskEdBox2.Text)
If day_cnt < 0 Then
   MsgBox "PLEASE ENTER VALID DATE RANGE"
   MaskEdBox2.Text = "__/__/____"
   MaskEdBox2.SetFocus
   Exit Sub
End If

If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False Then
    MsgBox "Please select between Ward/ICU/Patient/Unit"
    Exit Sub
End If

If Option1.Value = True Then 'Ward
    
    If cmb_ward.Text = "" Then
        MsgBox "Please select ward"
        Exit Sub
    End If
    
    If Opt_specimen.Value = True And cmb_specimen.Text = "" Then
        MsgBox "Please select specimen"
        Exit Sub
    End If
    
    '***** Added by shraddha on 09/12/09 *****
    If Opt_organism.Value = True And cmb_organism.Text = "" Then
        MsgBox "Please select organism"
        Exit Sub
    End If
    '*****************************************
    
    If Opt_antibiotic.Value = True And cmb_antibiotic.Text = "" Then
        MsgBox "Please select Antiboitic"
        Exit Sub
    End If
    
'    If Opt_antibiotic.Value = False And Opt_specimen.Value = False Then
'        MsgBox "Please select between specimen/antibiotic"
'        Exit Sub
'    End If
End If

If Option2.Value = True Then 'ICU
    If Opt_specimen.Value = True And cmb_specimen.Text = "" Then
        MsgBox "Please select specimen"
        Exit Sub
    End If
    
    '***** Added by shraddha on 09/12/09 *****
    If Opt_organism.Value = True And cmb_organism.Text = "" Then
        MsgBox "Please select organism"
        Exit Sub
    End If
    '*****************************************
    
    If Opt_antibiotic.Value = True And cmb_antibiotic.Text = "" Then
        MsgBox "Please select Antiboitic"
        Exit Sub
    End If
    
'    If Opt_antibiotic.Value = False And Opt_specimen.Value = False Then
'        MsgBox "Please select between specimen/antibiotic"
'        Exit Sub
'    End If
    
'InsDataPathogen
    
End If

If Option3.Value = True Then
    If MaskEdBox3.Text = "" Then
        MsgBox "Please enter Case No."
        Exit Sub
    End If
End If

If Option4.Value = True Then
    If cmb_unit.Text = "" Then
        MsgBox "Please select unit"
    End If
    If cmb_unit.Text <> "" And cmb_doctor.Visible = True Then
        If cmb_doctor.Text = "" Then
            MsgBox "Please select doctor"
        End If
    End If
End If
'**************************************************************
var_date_f = Year(MaskEdBox1.FormattedText) & "," & month(MaskEdBox1.FormattedText) & "," & Day(MaskEdBox1.FormattedText)
var_date_t = Year(MaskEdBox2.FormattedText) & "," & month(MaskEdBox2.FormattedText) & "," & Day(MaskEdBox2.FormattedText)
 
If Option1.Value = True Then 'Ward

    CR.ReportFileName = "C:\DISREPORTSEC\micro_icu1.rpt"
    CR.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)
    CR.Formulas(0) = "from=Date(" & var_date_f & ")"
    CR.Formulas(1) = "to=Date(" & var_date_t & ")"
    CR.Formulas(2) = "USRID = '" & UI & "'"
    CR.Formulas(3) = "ward = '" & wardname & "'"
    'CR.SelectionFormula = "{DMULPROCE.REQNDT}>=Date(" & var_date_f & ") AND {DMULPROCE.REQNDT}<=Date(" & var_date_t & ")"
    
    'CR.SelectionFormula = "{RREQUISITN.CASENO} = 'CE/06877' AND {rrequisitn.REQNDT}>=Date(" & var_date_f & ") AND {rrequisitn.REQNDT}<=Date(" & var_date_t & ")"
        
    'Only Ward
    If Option1.Value = True And Opt_antibiotic.Value = False And Opt_specimen.Value = False And Opt_organism.Value = False Then
        wardname = Trim(cmb_ward.Text)
        CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.WARD_NAME}= '" & wardname & "'"
    End If
    
    If Option1.Value = True And Opt_specimen.Value = True Then
        wardname = Trim(cmb_ward.Text)
        specimen = Trim(cmb_specimen.Text)
        If cmb_specimen.Text = "ALL" Then
            CR.Formulas(3) = "ward = '" & wardname & " (ALL Specimen)'"
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.WARD_NAME}= '" & wardname & "'"
        Else
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.WARD_NAME}= '" & wardname & "' and {PATHOGEN.SPECIMEN}='" & specimen & "'"
        End If
        
    End If
    
    If Opt_antibiotic.Value = True Then
        wardname = Trim(cmb_ward.Text)
        antibiotic = Trim(cmb_antibiotic.Text)
        If cmb_antibiotic.Text = "ALL" Then
            CR.Formulas(3) = "ward = '" & wardname & " (ALL Antibiotics S)'"
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.WARD_NAME}= '" & wardname & "' and ({PATHOGEN.ANTIBIO1}='S' or {PATHOGEN.ANTIBIO2}='S' or {PATHOGEN.ANTIBIO3}='S')"
        Else
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.WARD_NAME}= '" & wardname & "' and {PATHOGEN.ANTIBIOTIC}='" & antibiotic & "' and ({PATHOGEN.ANTIBIO1}='S' or {PATHOGEN.ANTIBIO2}='S' or {PATHOGEN.ANTIBIO3}='S')"
        End If
    End If
    
    '***** Shraddha 09/12/09 *****
    
    If Option1.Value = True And Opt_organism.Value = True Then
        wardname = Trim(cmb_ward.Text)
        organism = Trim(cmb_organism.Text)
        If cmb_organism.Text = "ALL" Then
            CR.Formulas(3) = "ward = '" & wardname & " (ALL Specimen)'"
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.WARD_NAME}= '" & wardname & "'"
        Else
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.WARD_NAME}= '" & wardname & "' and ({PATHOGEN.ORGSO1} like '*" & organism & "' or {PATHOGEN.ORGSO2} like '*" & organism & "' or {PATHOGEN.ORGSO3} like '*" & organism & "')"
        End If
        
    End If
    '*****************************
    
    CR.WindowState = crptMaximized
    CR.Action = 1

'**********************************
End If

If Option2.Value = True Then 'ICU

    CR.ReportFileName = "C:\DISREPORTSEC\micro_icu1.rpt"
    CR.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)
    CR.Formulas(0) = "from=Date(" & var_date_f & ")"
    CR.Formulas(1) = "to=Date(" & var_date_t & ")"
    CR.Formulas(2) = "USRID = '" & UI & "'"
    CR.Formulas(3) = "icu = 'ICU'"
            
    'Only ICU
    If Option2.Value = True And Opt_specimen.Value = False And Opt_antibiotic.Value = False And Opt_organism.Value = False Then
        CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.ICDS2}='YES'"
    End If
    
    If Opt_specimen.Value = True Then
        specimen = Trim(cmb_specimen.Text)
        If cmb_specimen = "ALL" Then
            CR.Formulas(3) = "icu = 'ICU' (All Specimen)"
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.ICDS2}='YES'"
        Else
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.ICDS2}='YES' and {PATHOGEN.SPECIMEN}='" & specimen & "'"
        End If
    End If
    
    If Opt_antibiotic.Value = True Then
        antibiotic = Trim(cmb_antibiotic.Text)
        If cmb_antibiotic = "ALL" Then
            CR.Formulas(3) = "icu = 'ICU' (All Antibiotics S) "
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.ICDS2}='YES' and ({PATHOGEN.ANTIBIO1}='S' or {PATHOGEN.ANTIBIO2}='S' or {PATHOGEN.ANTIBIO3}='S')"
        Else
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.ICDS2}='YES' and {PATHOGEN.ANTIBIOTIC}='" & antibiotic & "' and ({PATHOGEN.ANTIBIO1}='S' or {PATHOGEN.ANTIBIO2}='S' or {PATHOGEN.ANTIBIO3}='S')"
        End If
    End If
    
    '***** Shraddha 09/12/09 *****
    If Opt_organism.Value = True Then
        organism = Trim(cmb_organism.Text)
        If cmb_organism = "ALL" Then
            CR.Formulas(3) = "icu = 'ICU' (All Organism)"
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.ICDS2}='YES'"
        Else
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.ICDS2}='YES' and ({PATHOGEN.ORGSO1} like '*" & organism & "' or {PATHOGEN.ORGSO2} like '*" & organism & "' or {PATHOGEN.ORGSO3} like '*" & organism & "')"
        End If
    End If
    '*****************************
    
    CR.WindowState = crptMaximized
    'CR.Action = 1

'**********************************
End If

If Option3.Value = True Then 'Patient
c_no = MaskEdBox3.FormattedText
CR.ReportFileName = "C:\DISREPORTSEC\micro_icu1.rpt"
    CR.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)
    CR.Formulas(0) = "from=Date(" & var_date_f & ")"
    CR.Formulas(1) = "to=Date(" & var_date_t & ")"
    CR.Formulas(2) = "USRID = '" & UI & "'"
    CR.Formulas(3) = "caseno = ' : CASE No. " & c_no & "'" '05/08/09
    
    CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.caseno}= '" & c_no & "'"
    
    CR.WindowState = crptMaximized
    CR.Action = 1
End If

'**************** 05/08/2009 *******************
If Option4.Value = True Then 'ICU

    CR.ReportFileName = "C:\DISREPORTSEC\micro_icu1.rpt"
    CR.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)
    CR.Formulas(0) = "from=Date(" & var_date_f & ")"
    CR.Formulas(1) = "to=Date(" & var_date_t & ")"
    CR.Formulas(2) = "USRID = '" & UI & "'"
    'CR.Formulas(3) = "icu = 'ICU'"

    If cmb_doctor.Visible = True Then
        doctor = cmb_doctor.Text
        CR.Formulas(3) = "doc_uni = ': DOCTOR - " & doctor & "'"
        CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.refbydoc}='" & doctor & "'"
    Else
        unit = cmb_unit.Text
        
        If cmb_unit.Text = "ALL" Then
            CR.Formulas(3) = "doc_uni = ': UNIT - " & unit & "(All units)'"
            'CR.Formulas(3) = "doc_uni = ': UNIT - " & unit & "'"
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ")"
            
        Else
            CR.Formulas(3) = "doc_uni = ': UNIT - " & unit & "'"
            CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.refbydoc}='" & unit & "'"
        End If
    End If
    
    CR.WindowState = crptMaximized
    CR.Action = 1

End If
'*************************************************
End Sub

Private Sub MaskEdBox3_LostFocus()
If Option3.Value = True Then
    If MaskEdBox3.Text = "" Then
        MsgBox "Please enter Case No."
        Exit Sub
    End If
    MaskEdBox3.Text = UCase(padzero(5, "C", MaskEdBox3.Text, 3))
    'c_no = MaskEdBox3.FormattedText
End If
cmd_view.SetFocus
End Sub

Private Sub Opt_antibiotic_Click()
cmb_antibiotic.Visible = True
cmb_specimen.Visible = False
cmb_organism.Visible = False
'SendKeys "^{f4}"
cmb_antibiotic.SetFocus
End Sub

Private Sub opt_OP_Click()
cmb_ward.Visible = False
lbl_wards.Visible = False

'SendKeys "^{f4}"
'cmb_ward.SetFocus

MaskEdBox3.Visible = False
Shape_patient.Visible = False
lbl_caseno.Visible = False
MSFlexGrid1.Visible = False
Frame_colour.Visible = False
lbl_caption.Visible = False
'cmb_specimen.Visible = True
'lbl_specimen.Visible = True
'cmb_antibiotic.Visible = True
frame_SA.Visible = False
Shape_SA.Visible = False
Shape_WSA.Visible = False
cmb_unit.Visible = False
Shape_unit.Visible = False
lbl_unit.Visible = False
lbl_doctor.Visible = False
cmb_doctor.Visible = False
End Sub

Private Sub Opt_specimen_Click()
cmb_specimen.Visible = True
cmb_antibiotic.Visible = False
cmb_organism.Visible = False
'SendKeys "^{f4}"
cmb_specimen.SetFocus
End Sub

'***** Shraddha 08/12/09 *****
Private Sub Opt_organism_click()
cmb_organism.Visible = True
cmb_antibiotic.Visible = False
cmb_specimen.Visible = False
'SendKeys "^{f4}"
cmb_organism.SetFocus
End Sub
'*****************************

Private Sub Option1_Click()
cmb_ward.Visible = True
lbl_wards.Visible = True

'SendKeys "^{f4}"
cmb_ward.SetFocus

MaskEdBox3.Visible = False
Shape_patient.Visible = False
lbl_caseno.Visible = False
MSFlexGrid1.Visible = False
Frame_colour.Visible = False
lbl_caption.Visible = False
'cmb_specimen.Visible = True
'lbl_specimen.Visible = True
'cmb_antibiotic.Visible = True
frame_SA.Visible = True
Shape_SA.Visible = False
Shape_WSA.Visible = True
cmb_unit.Visible = False
Shape_unit.Visible = False
lbl_unit.Visible = False
lbl_doctor.Visible = False
cmb_doctor.Visible = False
End Sub

Private Sub Option2_Click()
frame_SA.Visible = True
Shape_SA.Visible = True
cmb_ward.Visible = False
lbl_wards.Visible = False
MaskEdBox3.Visible = False
Shape_patient.Visible = False
lbl_caseno.Visible = False
MSFlexGrid1.Visible = False
Frame_colour.Visible = False
lbl_caption.Visible = False
'cmb_specimen.Visible = True
'lbl_specimen.Visible = True
'cmb_antibiotic.Visible = True
Shape_WSA.Visible = False
cmb_unit.Visible = False
Shape_unit.Visible = False
lbl_unit.Visible = False
lbl_doctor.Visible = False
cmb_doctor.Visible = False
End Sub

Private Sub Option3_Click()
Shape_patient.Visible = True
MaskEdBox3.Visible = True
MaskEdBox3.Text = ""
MaskEdBox3.SetFocus
lbl_caseno.Visible = True
lbl_wards.Visible = False
cmb_ward.Visible = False
MSFlexGrid1.Visible = False
Frame_colour.Visible = False
lbl_caption.Visible = False
'cmb_specimen.Visible = False
'cmb_antibiotic.Visible = False
'lbl_specimen.Visible = False
frame_SA.Visible = False
Shape_SA.Visible = False
Shape_WSA.Visible = False
cmb_unit.Visible = False
Shape_unit.Visible = False
lbl_unit.Visible = False
lbl_doctor.Visible = False
cmb_doctor.Visible = False
End Sub

Private Sub Option4_Click()
cmb_unit.Visible = True
Shape_unit.Visible = True
lbl_unit.Visible = True
lbl_doctor.Visible = False
cmb_doctor.Visible = False
'SendKeys "^{f4}"
cmb_unit.SetFocus

MaskEdBox3.Visible = False
Shape_patient.Visible = False
lbl_caseno.Visible = False
lbl_wards.Visible = False
cmb_ward.Visible = False
MSFlexGrid1.Visible = False
Frame_colour.Visible = False
lbl_caption.Visible = False
'cmb_specimen.Visible = False
'cmb_antibiotic.Visible = False
'lbl_specimen.Visible = False
frame_SA.Visible = False
Shape_SA.Visible = False
Shape_WSA.Visible = False
End Sub

Private Sub cmd_view_Click()
'************************ Validations ******************************
If MaskEdBox1.Text = "__/__/____" Then
    MsgBox "Please enter From date "
    MaskEdBox1.Text = "__/__/____"
    MaskEdBox1.SetFocus
    Exit Sub
ElseIf IsDate(MaskEdBox1.Text) = False Then
    MsgBox "Please enter valid From date "
    MaskEdBox1.Text = "__/__/____"
    MaskEdBox1.SetFocus
    Exit Sub
End If

If MaskEdBox2.Text = "__/__/____" Then
    MsgBox "Please enter To date "
    MaskEdBox2.Text = "__/__/____"
    MaskEdBox2.SetFocus
    Exit Sub
ElseIf IsDate(MaskEdBox2.Text) = False Then
    MsgBox "Please enter valid To date "
    MaskEdBox2.Text = "__/__/____"
    MaskEdBox2.SetFocus
    Exit Sub
End If

Dim day_cnt
day_cnt = DateDiff("d", MaskEdBox1.Text, MaskEdBox2.Text)
If day_cnt < 0 Then
   MsgBox "PLEASE ENTER VALID DATE RANGE"
   MaskEdBox2.Text = "__/__/____"
   MaskEdBox2.SetFocus
   Exit Sub
End If

If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False And opt_OP.Value = False Then
    MsgBox "Please select between Ward/ICU/Patient/Unit"
    Exit Sub
End If

If Option1.Value = True Then 'Ward
    
    If cmb_ward.Text = "" Then
        MsgBox "Please select ward"
        Exit Sub
    End If
    
    If Opt_specimen.Value = True And cmb_specimen.Text = "" Then
        MsgBox "Please select specimen"
        Exit Sub
    End If
    
    '***** Shraddha 09/12/09 *****
    If Opt_organism.Value = True And cmb_organism.Text = "" Then
        MsgBox "Please select organism"
        Exit Sub
    End If
    '*****************************
    
    If Opt_antibiotic.Value = True And cmb_antibiotic.Text = "" Then
        MsgBox "Please select Antiboitic"
        Exit Sub
    End If
    
'    If Opt_antibiotic.Value = False And Opt_specimen.Value = False Then
'        MsgBox "Please select between specimen/antibiotic"
'        Exit Sub
'    End If
End If

If Option2.Value = True Then 'ICU
    If Opt_specimen.Value = True And cmb_specimen.Text = "" Then
        MsgBox "Please select specimen"
        Exit Sub
    End If
    
    '***** Shraddha 09/12/09 *****
    If Opt_organism.Value = True And cmb_organism.Text = "" Then
        MsgBox "Please select organism"
        Exit Sub
    End If
    '*****************************
    
    If Opt_antibiotic.Value = True And cmb_antibiotic.Text = "" Then
        MsgBox "Please select Antiboitic"
        Exit Sub
    End If
    
'    If Opt_antibiotic.Value = False And Opt_specimen.Value = False Then
'        MsgBox "Please select between specimen/antibiotic"
'        Exit Sub
'    End If
End If

If Option3.Value = True Then
    If MaskEdBox3.Text = "" Then
        MsgBox "Please enter Case No."
        Exit Sub
    End If
End If

If Option4.Value = True Then
    If cmb_unit.Text = "" Then
        MsgBox "Please select unit"
    End If
    If cmb_unit.Text <> "" And cmb_doctor.Visible = True Then
        If cmb_doctor.Text = "" Then
            MsgBox "Please select doctor"
        End If
    End If
End If
'*******************************************************************
MSFlexGrid1.clear '05/08/2009
'Frame_colour.Visible = True
MousePointer = vbHourglass
n = 1
var_date_f = Year(MaskEdBox1.FormattedText) & "," & month(MaskEdBox1.FormattedText) & "," & Day(MaskEdBox1.FormattedText)
var_date_t = Year(MaskEdBox2.FormattedText) & "," & month(MaskEdBox2.FormattedText) & "," & Day(MaskEdBox2.FormattedText)
 
MSFlexGrid1.Visible = True
MSFlexGrid1.Rows = 2
MSFlexGrid1.Cols = 15

MSFlexGrid1.TextMatrix(0, 0) = "SNO."
MSFlexGrid1.TextMatrix(0, 1) = "CASE NO."
MSFlexGrid1.TextMatrix(0, 2) = "LAB NO."
MSFlexGrid1.TextMatrix(0, 3) = "REQUISITION NO."
MSFlexGrid1.TextMatrix(0, 4) = "NAME"
MSFlexGrid1.TextMatrix(0, 5) = "SEX"
MSFlexGrid1.TextMatrix(0, 6) = "AGE" 'Shraddha 29/07/09
MSFlexGrid1.TextMatrix(0, 7) = "SPECIMEN"
MSFlexGrid1.TextMatrix(0, 8) = "COLLECTION DATE" 'Added on 29/07/09 by Shraddha
MSFlexGrid1.TextMatrix(0, 9) = "ORGANISM1"
MSFlexGrid1.TextMatrix(0, 10) = "ANTIBIOTIC1"
MSFlexGrid1.TextMatrix(0, 11) = "ORGANISM2"
MSFlexGrid1.TextMatrix(0, 12) = "ANTIBIOTIC2"
MSFlexGrid1.TextMatrix(0, 13) = "ORGANISM3"
MSFlexGrid1.TextMatrix(0, 14) = "ANTIBIOTIC3"
MSFlexGrid1.RowHeight(0) = 350

If Option1.Value = True Then 'Ward

'********** According to ward and specimen ********
'sql_pathogen_ward = "select * from RISLIB.pathogen where WARD_NAME=? AND SPECIMEN=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
'Set sel_pathogen_ward_qr = rdb.CreateQuery("", sql_pathogen_ward)
'**************************************************

'********** According to ward and Antibiotic 15/06/09 ********
'sql_pathogen_ward = "select * from RISLIB.pathogen where WARD_NAME=? AND ANTIBIOTIC=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
'Set sel_pathogen_ward_qr = rdb.CreateQuery("", sql_pathogen_ward)
'*************************************************************
    
    '***********************
    If Opt_specimen.Value = True Then
        wardname = Trim(cmb_ward.Text)
        specimen = Trim(cmb_specimen.Text)
        If specimen = "ALL" Then
            sql_pathogen_ward = "select * from RISLIB.pathogen_new where WARD_NAME=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
            Set sel_pathogen_ward_qr = rdb.CreateQuery("", sql_pathogen_ward)
                    
            sel_pathogen_ward_qr(0) = wardname
            'sel_pathogen_ward_qr(1) = specimen
            sel_pathogen_ward_qr(1) = var_date_f
            sel_pathogen_ward_qr(2) = var_date_t
            Set sel_pathogen_ward_rs = sel_pathogen_ward_qr.OpenResultset(1, 3)
        Else
            sql_pathogen_ward = "select * from RISLIB.pathogen_new where WARD_NAME=? AND SPECIMEN=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
            Set sel_pathogen_ward_qr = rdb.CreateQuery("", sql_pathogen_ward)
                    
            sel_pathogen_ward_qr(0) = wardname
            sel_pathogen_ward_qr(1) = specimen
            sel_pathogen_ward_qr(2) = var_date_f
            sel_pathogen_ward_qr(3) = var_date_t
            Set sel_pathogen_ward_rs = sel_pathogen_ward_qr.OpenResultset(1, 3)
        End If
    End If
    
    If Opt_antibiotic.Value = True Then
        wardname = Trim(cmb_ward.Text)
        antibiotic = Trim(cmb_antibiotic.Text)
        If antibiotic = "ALL" Then
            sql_pathogen_ward = "select * from RISLIB.pathogen_new where WARD_NAME=? and (ANTIBIO1='S' or ANTIBIO2='S' or ANTIBIO3='S') AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
            Set sel_pathogen_ward_qr = rdb.CreateQuery("", sql_pathogen_ward)
        
            sel_pathogen_ward_qr(0) = wardname
            'sel_pathogen_ward_qr(1) = antibiotic
            sel_pathogen_ward_qr(1) = var_date_f
            sel_pathogen_ward_qr(2) = var_date_t
            Set sel_pathogen_ward_rs = sel_pathogen_ward_qr.OpenResultset(1, 3)

        Else
            sql_pathogen_ward = "select * from RISLIB.pathogen_new where WARD_NAME=? AND ANTIBIOTIC=? and (ANTIBIO1='S' or ANTIBIO2='S' or ANTIBIO3='S') AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
            Set sel_pathogen_ward_qr = rdb.CreateQuery("", sql_pathogen_ward)
            
            sel_pathogen_ward_qr(0) = wardname
            sel_pathogen_ward_qr(1) = antibiotic
            sel_pathogen_ward_qr(2) = var_date_f
            sel_pathogen_ward_qr(3) = var_date_t
            Set sel_pathogen_ward_rs = sel_pathogen_ward_qr.OpenResultset(1, 3)
        End If
    End If
    
    '***** Shraddha 09/12/09 *****
    If Opt_organism.Value = True Then
        wardname = Trim(cmb_ward.Text)
        organism = Trim(cmb_organism.Text)
        If organism = "ALL" Then
            sql_pathogen_ward = "select * from RISLIB.pathogen_new where WARD_NAME=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
            Set sel_pathogen_ward_qr = rdb.CreateQuery("", sql_pathogen_ward)
        
            sel_pathogen_ward_qr(0) = wardname
            'sel_pathogen_ward_qr(1) = antibiotic
            sel_pathogen_ward_qr(1) = var_date_f
            sel_pathogen_ward_qr(2) = var_date_t
            Set sel_pathogen_ward_rs = sel_pathogen_ward_qr.OpenResultset(1, 3)

        Else
            sql_pathogen_ward = "select * from RISLIB.pathogen_new where WARD_NAME=? and (ORGSO1 like ? or ORGSO2 like ? or ORGSO3 like ?) AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
            Set sel_pathogen_ward_qr = rdb.CreateQuery("", sql_pathogen_ward)
            
            sel_pathogen_ward_qr(0) = wardname
            
            sel_pathogen_ward_qr(1) = "%" & organism
            sel_pathogen_ward_qr(2) = "%" & organism
            sel_pathogen_ward_qr(3) = "%" & organism
            
            sel_pathogen_ward_qr(4) = var_date_f
            sel_pathogen_ward_qr(5) = var_date_t
            Set sel_pathogen_ward_rs = sel_pathogen_ward_qr.OpenResultset(1, 3)
        End If
    End If
    '-----------------------------
    '*****************************
    
    '******* Only according to ward 29/06/09 **********
    'sql_pathogen_ward = "select * from RISLIB.pathogen where WARD_NAME=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
    If Option1.Value = True And Opt_antibiotic.Value = False And Opt_specimen.Value = False And Opt_organism.Value = False Then
        wardname = Trim(cmb_ward.Text)
        'antibiotic = Trim(cmb_antibiotic.Text)
        
         '***** Shraddha 12/11/10 *****
        If wardname = "ALL" Then
    
'            sql_pathogen_ward = "select * from RISLIB.pathogen where REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
'            Set sel_pathogen_Allward_qr = rdb.CreateQuery("", sql_pathogen_ward)
'
'            sel_pathogen_Allward_qr.QueryTimeout = 4000
'            sel_pathogen_Allward_qr(0) = var_date_f
'            sel_pathogen_Allward_qr(1) = var_date_t
'            Set sel_pathogen_Allward_rs = sel_pathogen_Allward_qr.OpenResultset(1, 3)

            sql_pathogen_ward = "select * from RISLIB.pathogen_new where REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
            Set sel_pathogen_ward_qr = rdb.CreateQuery("", sql_pathogen_ward)
        
            sel_pathogen_ward_qr.QueryTimeout = 4000
            sel_pathogen_ward_qr(0) = var_date_f
            sel_pathogen_ward_qr(1) = var_date_t
            Set sel_pathogen_ward_rs = sel_pathogen_ward_qr.OpenResultset(1, 3)
                           
        Else
        '*****************************
        
            sql_pathogen_ward = "select * from RISLIB.pathogen_new where WARD_NAME=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO ,FIRST_NAME, MID_NAME "
            Set sel_pathogen_ward_qr = rdb.CreateQuery("", sql_pathogen_ward)
            
            sel_pathogen_ward_qr(0) = wardname
            'sel_pathogen_ward_qr(1) = antibiotic
            sel_pathogen_ward_qr(1) = var_date_f
            sel_pathogen_ward_qr(2) = var_date_t
            Set sel_pathogen_ward_rs = sel_pathogen_ward_qr.OpenResultset(1, 3)
        
            i = sel_pathogen_ward_rs.RowCount
            
        End If
    End If
    '**************************************************
    
   
    
    i = sel_pathogen_ward_rs.RowCount
    If i = 0 Then
        title_chk = True
    Else
        title_chk = False
    End If
    MSFlexGrid1.Rows = sel_pathogen_ward_rs.RowCount + 1
    For i = 1 To sel_pathogen_ward_rs.RowCount
        MSFlexGrid1.ColWidth(0) = 500
        MSFlexGrid1.ColWidth(1) = 1000
        MSFlexGrid1.ColWidth(2) = 900 'labno
        MSFlexGrid1.ColWidth(3) = 1500
        MSFlexGrid1.ColWidth(4) = 2400
        MSFlexGrid1.ColWidth(5) = 500 'Gender
        MSFlexGrid1.ColWidth(6) = 500 'Age
        MSFlexGrid1.ColWidth(7) = 2000
        MSFlexGrid1.ColWidth(8) = 2000 'Collection Date
        MSFlexGrid1.ColWidth(9) = 2400
        MSFlexGrid1.ColWidth(10) = 2400
        MSFlexGrid1.ColWidth(11) = 2400
        MSFlexGrid1.ColWidth(12) = 2400
        MSFlexGrid1.ColWidth(13) = 2400
        MSFlexGrid1.ColWidth(14) = 1900
        MSFlexGrid1.RowHeight(i) = 350
'''''''''''''''''''''''''''''''''''''''

        If i <> 1 Then
            If Trim(cno) = Trim(sel_pathogen_ward_rs("caseno")) Then
                MSFlexGrid1.TextMatrix(i, 1) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 1) = sel_pathogen_ward_rs("caseno")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 1) = sel_pathogen_ward_rs("caseno")
        End If
        'cno = sel_pathogen_ward_rs("caseno")
        
        If MSFlexGrid1.TextMatrix(i, 1) <> "" Then
            MSFlexGrid1.TextMatrix(i, 0) = n
            n = n + 1
        End If
        
        '*********** Added by Shraddha on 12/06/09 ***********
        sel_labno_qr(0) = Trim(sel_pathogen_ward_rs("reqnno"))
        Set sel_labno_rs = sel_labno_qr.OpenResultset(1, 3)
         
        If i <> 1 Then
            If Trim(lab_no) = Trim(sel_pathogen_ward_rs("pathno")) And cno = sel_pathogen_ward_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 2) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 2) = Trim(sel_pathogen_ward_rs("pathno"))
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 2) = Trim(sel_pathogen_ward_rs("pathno"))
        End If
        lab_no = Trim(sel_pathogen_ward_rs("pathno"))
        '*****************************************************
    
        If i <> 1 Then
            If Trim(rno) = Trim(sel_pathogen_ward_rs("reqnno")) And cno = sel_pathogen_ward_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 3) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 3) = sel_pathogen_ward_rs("reqnno")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 3) = sel_pathogen_ward_rs("reqnno")
        End If
        rno = sel_pathogen_ward_rs("reqnno")
       
        
        If i <> 1 Then
            'If Trim(pname) = Trim(Trim(sel_pathogen_ward_rs("tag")) & " " & Trim(sel_pathogen_ward_rs("FIRST_NAME")) & " " & Trim(sel_pathogen_ward_rs("MID_NAME")) & " " & Trim(sel_pathogen_ward_rs("LAST_NAME"))) Then
            If cno = sel_pathogen_ward_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 4) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 4) = Trim(sel_pathogen_ward_rs("tag")) & " " & Trim(sel_pathogen_ward_rs("FIRST_NAME")) & " " & Trim(sel_pathogen_ward_rs("MID_NAME")) & " " & Trim(sel_pathogen_ward_rs("LAST_NAME"))
            End If
        Else
            MSFlexGrid1.TextMatrix(i, 4) = Trim(sel_pathogen_ward_rs("tag")) & " " & Trim(sel_pathogen_ward_rs("FIRST_NAME")) & " " & Trim(sel_pathogen_ward_rs("MID_NAME")) & " " & Trim(sel_pathogen_ward_rs("LAST_NAME"))
        End If

        If sel_pathogen_ward_rs("tag") <> "" Then
            sal = sel_pathogen_ward_rs("tag")
        End If
        If sel_pathogen_ward_rs("FIRST_NAME") <> "" Then
            FNAME = sel_pathogen_ward_rs("FIRST_NAME")
        End If
        If sel_pathogen_ward_rs("MID_NAME") <> "" Then
            mname = sel_pathogen_ward_rs("MID_NAME")
        End If
        If sel_pathogen_ward_rs("LAST_NAME") <> "" Then
            lname = sel_pathogen_ward_rs("LAST_NAME")
        End If
        pname = Trim(sal) & " " & Trim(FNAME) & " " & Trim(mname) & " " & Trim(lname)
     
  
        '****************** Added on 08/06/09 ********************
        If Len(Trim(sel_pathogen_ward_rs("BIRTH_DATE"))) <> 0 Then
            age = DateDiff("YYYY", sel_pathogen_ward_rs("BIRTH_DATE"), CStr(DATE11))
        End If
        '*********************************************************
          
        'If I <> 1 Then
        '    If Trim(gender) = Trim(sel_pathogen_ward_rs("SEX")) & "/" & age Then
        '        MSFlexGrid1.TextMatrix(I, 5) = ""
        '    Else
        '        MSFlexGrid1.TextMatrix(I, 5) = sel_pathogen_ward_rs("sex") & "/" & age
        '    End If
        'Else
        '        MSFlexGrid1.TextMatrix(I, 5) = sel_pathogen_ward_rs("sex") & "/" & age
        'End If
        'gender = sel_pathogen_ward_rs("sex") & "/" & age

        '************ Shraddha 13/06/09 ************
'        If MSFlexGrid1.TextMatrix(I, 1) <> "" Then
'            'MSFlexGrid1.TextMatrix(I, 5) = Trim(sel_pathogen_ward_rs("SEX")) & "/" & age
'            MSFlexGrid1.TextMatrix(I, 5) = Trim(sel_pathogen_ward_rs("SEX"))
'        End If
        'Once occuring
        If i <> 1 Then
            If MSFlexGrid1.TextMatrix(i, 1) = "" Then
                MSFlexGrid1.TextMatrix(i, 5) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 5) = Trim(sel_pathogen_ward_rs("SEX"))
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 5) = Trim(sel_pathogen_ward_rs("SEX"))
        End If
        
        '*******************************************
        
        '******* 29/07/09 *******
'        If MSFlexGrid1.TextMatrix(I, 1) <> "" Then
'            MSFlexGrid1.TextMatrix(I, 6) = age
'        End If
'Once occuring
        If i <> 1 Then
            If MSFlexGrid1.TextMatrix(i, 1) = "" Then
                MSFlexGrid1.TextMatrix(i, 6) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 6) = age
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 6) = age
        End If
        '***********************

        'If I <> 1 Then
        '    If Trim(nmaterial) = Trim(sel_pathogen_ward_rs("SPECIMEN")) Then
        '        MSFlexGrid1.TextMatrix(I, 6) = ""
        '    Else
        '        MSFlexGrid1.TextMatrix(I, 6) = sel_pathogen_ward_rs("SPECIMEN")
        '    End If
        'Else
        '    MSFlexGrid1.TextMatrix(I, 6) = sel_pathogen_ward_rs("SPECIMEN")
        'End If
        'nmaterial = sel_pathogen_ward_rs("SPECIMEN")
        
        '************ Shraddha 13/06/09 ************
'        If MSFlexGrid1.TextMatrix(I, 1) <> "" Then
'            MSFlexGrid1.TextMatrix(I, 7) = Trim(sel_pathogen_ward_rs("SPECIMEN"))
'        End If
        
        '***** Shraddha 10/10/2011 *****
        If Trim(sel_pathogen_ward_rs("ORGSO1")) <> "" Or Trim(sel_pathogen_ward_rs("ORGSO2")) <> "" Or Trim(sel_pathogen_ward_rs("ORGSO3")) <> "" Then
            patho_grown = "  Pathogen grown : Yes"
        Else
            patho_grown = "  Pathogen grown : No"
        End If
        '*******************************
        
        If i <> 1 Then
            If Trim(UCase(SP)) = Trim(UCase(sel_pathogen_ward_rs("SPECIMEN"))) And cno = sel_pathogen_ward_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 7) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 7) = Trim(sel_pathogen_ward_rs("SPECIMEN")) & patho_grown
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 7) = Trim(sel_pathogen_ward_rs("SPECIMEN")) & patho_grown
        End If
        SP = Trim(sel_pathogen_ward_rs("SPECIMEN"))
        '*******************************************

        '*********** Shraddha Pandit 29/07/09 **************
        'Collection Date 8
'        If Trim(sel_pathogen_ward_rs("samcoldate")) <> "" Then
'
'            MSFlexGrid1.TextMatrix(I, 8) = Format(sel_pathogen_ward_rs("samcoldate"), "dd/mm/yyyy")
'
'            colldt = Format(sel_pathogen_ward_rs("samcoldate"), "dd/mm/yyyy")
'        End If
        
        If i <> 1 Then
            If Trim(colldt) = Format(sel_pathogen_ward_rs("samcoldate"), "dd/mm/yyyy") And cno = sel_pathogen_ward_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 8) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 8) = Format(sel_pathogen_ward_rs("samcoldate"), "dd/mm/yyyy")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 8) = Format(sel_pathogen_ward_rs("samcoldate"), "dd/mm/yyyy")
        End If
        If (sel_pathogen_ward_rs("samcoldate")) <> Null Or (sel_pathogen_ward_rs("samcoldate")) <> "" Then
            colldt = Format(sel_pathogen_ward_rs("samcoldate"), "dd/mm/yyyy")
        End If
        '***************************************************
     
        If Trim(sel_pathogen_ward_rs("ORGSO1")) <> "" Then
            If i <> 1 Then
                'If Trim(or1) = Trim(sel_pathogen_ward_rs("ORGSO1")) Then
                If Trim(or1) = Trim(sel_pathogen_ward_rs("ORGSO1")) And cno = sel_pathogen_ward_rs("caseno") Then
                
                'If MSFlexGrid1.TextMatrix(I, 1) = "" Then
                    MSFlexGrid1.TextMatrix(i, 9) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 9) = sel_pathogen_ward_rs("ORGSO1")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 9) = sel_pathogen_ward_rs("ORGSO1")
            End If
            or1 = sel_pathogen_ward_rs("ORGSO1")
        End If
        

        If Trim(sel_pathogen_ward_rs("ORGSO1")) <> "" Then
            If Trim(sel_pathogen_ward_rs("ANTIBIO1")) <> "" Then
                If i <> 1 Then
                    'If Trim(antbt1) = Trim(sel_pathogen_ward_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO1")) & ")" Then
                    If Trim(antbt1) = Trim(sel_pathogen_ward_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO1")) & ")" And cno = sel_pathogen_ward_rs("caseno") Then
                    
                        MSFlexGrid1.TextMatrix(i, 10) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 10) = sel_pathogen_ward_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO1")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 10) = sel_pathogen_ward_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO1")) & ")"
                End If
                antbt1 = Trim(sel_pathogen_ward_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO1")) & ")"
                
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 10
                If Trim(sel_pathogen_ward_rs("ANTIBIO1")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80 'Green
                End If
                If Trim(sel_pathogen_ward_rs("ANTIBIO1")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF 'Yellow
                End If
                If Trim(sel_pathogen_ward_rs("ANTIBIO1")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF 'Red
                End If
                If Opt_antibiotic.Value = True Then
                    MSFlexGrid1.TextMatrix(i, 10) = Trim(sel_pathogen_ward_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO1")) & ")"
                End If
            End If
        End If
        
        If Trim(sel_pathogen_ward_rs("ORGSO2")) <> "" Then
            If i <> 1 Then
                If Trim(or2) = Trim(sel_pathogen_ward_rs("ORGSO2")) And cno = sel_pathogen_ward_rs("caseno") Then
                    MSFlexGrid1.TextMatrix(i, 11) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 11) = sel_pathogen_ward_rs("ORGSO2")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 11) = sel_pathogen_ward_rs("ORGSO2")
            End If
            or2 = sel_pathogen_ward_rs("ORGSO2")
        End If


        If Trim(sel_pathogen_ward_rs("ORGSO2")) <> "" Then
            If Trim(sel_pathogen_ward_rs("ANTIBIO2")) <> "" Then
                If i <> 1 Then
                    If Trim(antbt2) = Trim(sel_pathogen_ward_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO2")) & ")" And cno = sel_pathogen_ward_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 12) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 12) = sel_pathogen_ward_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO2")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 12) = sel_pathogen_ward_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO2")) & ")"
                End If
                antbt2 = Trim(sel_pathogen_ward_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO2")) & ")"
                
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 12
                If Trim(sel_pathogen_ward_rs("ANTIBIO2")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_ward_rs("ANTIBIO2")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_ward_rs("ANTIBIO2")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
                If Opt_antibiotic.Value = True Then
                    MSFlexGrid1.TextMatrix(i, 12) = Trim(sel_pathogen_ward_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO2")) & ")"
                End If
            End If
        End If
        
        If Trim(sel_pathogen_ward_rs("ORGSO3")) <> "" Then
            If i <> 1 Then
                If Trim(or3) = Trim(sel_pathogen_ward_rs("ORGSO3")) And cno = sel_pathogen_ward_rs("caseno") Then
                    MSFlexGrid1.TextMatrix(i, 13) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 13) = sel_pathogen_ward_rs("ORGSO3")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 13) = sel_pathogen_ward_rs("ORGSO3")
            End If
            or3 = sel_pathogen_ward_rs("ORGSO3")
        End If


        If Trim(sel_pathogen_ward_rs("ORGSO3")) <> "" Then
            If Trim(sel_pathogen_ward_rs("ANTIBIO3")) <> "" Then
                If i <> 1 Then
                    If Trim(antbt3) = Trim(sel_pathogen_ward_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO3")) & ")" And cno = sel_pathogen_ward_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 14) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 14) = sel_pathogen_ward_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO3")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 14) = sel_pathogen_ward_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO3")) & ")"
                End If
                antbt3 = Trim(sel_pathogen_ward_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO3")) & ")"
                
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 14
                If Trim(sel_pathogen_ward_rs("ANTIBIO3")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_ward_rs("ANTIBIO3")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_ward_rs("ANTIBIO3")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
                If Opt_antibiotic.Value = True Then
                    MSFlexGrid1.TextMatrix(i, 14) = Trim(sel_pathogen_ward_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_ward_rs("ANTIBIO3")) & ")"
                End If
            End If
        End If
        cno = sel_pathogen_ward_rs("caseno")
        sel_pathogen_ward_rs.MoveNext
    Next

End If
'********************* ward ends ************************

If Option2.Value = True Then '********* ICU begins *************
        
    If Opt_specimen.Value = True Then
        specimen = Trim(cmb_specimen.Text)
        If specimen = "ALL" Then
            sql_pathogen_icu = "select * from RISLIB.pathogen_new where ICDS2='YES' AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
            Set sel_pathogen_icu_qr = rdb.CreateQuery("", sql_pathogen_icu)
            
            'sel_pathogen_icu_qr(0) = specimen
            sel_pathogen_icu_qr(0) = var_date_f
            sel_pathogen_icu_qr(1) = var_date_t
            Set sel_pathogen_icu_rs = sel_pathogen_icu_qr.OpenResultset(1, 3)
        Else
            sql_pathogen_icu = "select * from RISLIB.pathogen_new where ICDS2='YES' AND SPECIMEN=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
            Set sel_pathogen_icu_qr = rdb.CreateQuery("", sql_pathogen_icu)
            specimen = Trim(cmb_specimen.Text)
            sel_pathogen_icu_qr(0) = specimen
            sel_pathogen_icu_qr(1) = var_date_f
            sel_pathogen_icu_qr(2) = var_date_t
            Set sel_pathogen_icu_rs = sel_pathogen_icu_qr.OpenResultset(1, 3)
        End If
    End If
    
    '***** Shraddha 09/12/09 *****
    If Opt_organism.Value = True Then
        organism = Trim(cmb_organism.Text)
        If organism = "ALL" Then
            sql_pathogen_icu = "select * from RISLIB.pathogen_new where ICDS2='YES' AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
            Set sel_pathogen_icu_qr = rdb.CreateQuery("", sql_pathogen_icu)
            
            'sel_pathogen_icu_qr(0) = specimen
            sel_pathogen_icu_qr(0) = var_date_f
            sel_pathogen_icu_qr(1) = var_date_t
            Set sel_pathogen_icu_rs = sel_pathogen_icu_qr.OpenResultset(1, 3)
        Else
            sql_pathogen_icu = "select * from RISLIB.pathogen_new where ICDS2='YES' AND (ORGSO1 like ? or ORGSO2 like ? or ORGSO3 like ?) AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
            Set sel_pathogen_icu_qr = rdb.CreateQuery("", sql_pathogen_icu)
            specimen = Trim(cmb_specimen.Text)
            sel_pathogen_icu_qr(0) = "%" & organism
            sel_pathogen_icu_qr(1) = "%" & organism
            sel_pathogen_icu_qr(2) = "%" & organism
            
            sel_pathogen_icu_qr(3) = var_date_f
            sel_pathogen_icu_qr(4) = var_date_t
            Set sel_pathogen_icu_rs = sel_pathogen_icu_qr.OpenResultset(1, 3)
        End If
    End If
    '*****************************
    
    If Opt_antibiotic.Value = True Then
        antibiotic = Trim(cmb_antibiotic.Text)
        If antibiotic = "ALL" Then
            sql_pathogen_icu = "select * from RISLIB.pathogen_new where ICDS2='YES' AND (ANTIBIO1='S' or ANTIBIO2='S' or ANTIBIO3='S') AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
            Set sel_pathogen_icu_qr = rdb.CreateQuery("", sql_pathogen_icu)
                        
            sel_pathogen_icu_qr(0) = var_date_f
            sel_pathogen_icu_qr(1) = var_date_t
            Set sel_pathogen_icu_rs = sel_pathogen_icu_qr.OpenResultset(1, 3)
        Else
            sql_pathogen_icu = "select * from RISLIB.pathogen_new where ICDS2='YES' AND ANTIBIOTIC=? and (ANTIBIO1='S' or ANTIBIO2='S' or ANTIBIO3='S') AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
            Set sel_pathogen_icu_qr = rdb.CreateQuery("", sql_pathogen_icu)
            antibiotic = Trim(cmb_antibiotic.Text)
            sel_pathogen_icu_qr(0) = antibiotic
            sel_pathogen_icu_qr(1) = var_date_f
            sel_pathogen_icu_qr(2) = var_date_t
            Set sel_pathogen_icu_rs = sel_pathogen_icu_qr.OpenResultset(1, 3)
        End If
    End If
    
    '******* Only according to ICU ***********
    'sql_pathogen_icu = "select * from RISLIB.pathogen where ICDS2='YES' AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
    If Option2.Value = True And Opt_specimen.Value = False And Opt_antibiotic.Value = False And Opt_organism.Value = False Then
        sql_pathogen_icu = "select * from RISLIB.pathogen_new where ICDS2='YES' AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
        Set sel_pathogen_icu_qr = rdb.CreateQuery("", sql_pathogen_icu)
        'antibiotic = Trim(cmb_antibiotic.Text)
        'sel_pathogen_icu_qr(0) = antibiotic
        sel_pathogen_icu_qr(0) = var_date_f
        sel_pathogen_icu_qr(1) = var_date_t
        Set sel_pathogen_icu_rs = sel_pathogen_icu_qr.OpenResultset(1, 3)
    End If
    '*****************************************
    
    i = sel_pathogen_icu_rs.RowCount
    If i = 0 Then
        title_chk = True
    Else
        title_chk = False
    End If
    
    MSFlexGrid1.Rows = sel_pathogen_icu_rs.RowCount + 1
    For i = 1 To sel_pathogen_icu_rs.RowCount
        MSFlexGrid1.ColWidth(0) = 500
        MSFlexGrid1.ColWidth(1) = 1000
        MSFlexGrid1.ColWidth(2) = 900  'labno
        MSFlexGrid1.ColWidth(3) = 1500
        MSFlexGrid1.ColWidth(4) = 2400
        MSFlexGrid1.ColWidth(5) = 500 'Gender
        MSFlexGrid1.ColWidth(6) = 500 'age
        MSFlexGrid1.ColWidth(7) = 2000
        MSFlexGrid1.ColWidth(8) = 2000 'Collection Date
        MSFlexGrid1.ColWidth(9) = 2400
        MSFlexGrid1.ColWidth(10) = 2400
        MSFlexGrid1.ColWidth(11) = 2400
        MSFlexGrid1.ColWidth(12) = 2400
        MSFlexGrid1.ColWidth(13) = 2400
        MSFlexGrid1.ColWidth(14) = 1900
        MSFlexGrid1.RowHeight(i) = 350
        '''''''''''''''''''''''''''''''''''''''
        
        If i <> 1 Then
            If Trim(cno) = Trim(sel_pathogen_icu_rs("caseno")) Then
                MSFlexGrid1.TextMatrix(i, 1) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 1) = sel_pathogen_icu_rs("caseno")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 1) = sel_pathogen_icu_rs("caseno")
        End If
        'cno = sel_pathogen_icu_rs("caseno")
        
        If MSFlexGrid1.TextMatrix(i, 1) <> "" Then
            MSFlexGrid1.TextMatrix(i, 0) = n
            n = n + 1
        End If
        
        '*********** Added by Shraddha on 12/06/09 ***********
        sel_labno_qr(0) = Trim(sel_pathogen_icu_rs("reqnno"))
        Set sel_labno_rs = sel_labno_qr.OpenResultset(1, 3)
         
        If i <> 1 Then
            'If Trim(lab_no) = Trim(sel_pathogen_icu_rs("pathno")) Then
            If Trim(lab_no) = Trim(sel_pathogen_icu_rs("pathno")) And cno = sel_pathogen_icu_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 2) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 2) = Trim(sel_pathogen_icu_rs("pathno"))
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 2) = Trim(sel_pathogen_icu_rs("pathno"))
        End If
        lab_no = Trim(sel_pathogen_icu_rs("pathno"))
        '*****************************************************
      
        If i <> 1 Then
            'If Trim(rno) = Trim(sel_pathogen_icu_rs("reqnno")) Then
            If Trim(rno) = Trim(sel_pathogen_icu_rs("reqnno")) And cno = sel_pathogen_icu_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 3) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 3) = Trim(sel_pathogen_icu_rs("reqnno"))
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 3) = Trim(sel_pathogen_icu_rs("reqnno"))
        End If
        rno = sel_pathogen_icu_rs("reqnno")
        
        
        If i <> 1 Then
            'If Trim(pname) = Trim(Trim(sel_pathogen_icu_rs("tag")) & " " & Trim(sel_pathogen_icu_rs("FIRST_NAME")) & " " & Trim(sel_pathogen_icu_rs("MID_NAME")) & " " & Trim(sel_pathogen_icu_rs("LAST_NAME"))) Then
            If cno = sel_pathogen_icu_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 4) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 4) = Trim(sel_pathogen_icu_rs("tag")) & " " & Trim(sel_pathogen_icu_rs("FIRST_NAME")) & " " & Trim(sel_pathogen_icu_rs("MID_NAME")) & " " & Trim(sel_pathogen_icu_rs("LAST_NAME"))
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 4) = Trim(sel_pathogen_icu_rs("tag")) & " " & Trim(sel_pathogen_icu_rs("FIRST_NAME")) & " " & Trim(sel_pathogen_icu_rs("MID_NAME")) & " " & Trim(sel_pathogen_icu_rs("LAST_NAME"))
        End If

        If sel_pathogen_icu_rs("tag") <> "" Then
            sal = sel_pathogen_icu_rs("tag")
        End If
        If sel_pathogen_icu_rs("FIRST_NAME") <> "" Then
            FNAME = sel_pathogen_icu_rs("FIRST_NAME")
        End If
        If sel_pathogen_icu_rs("MID_NAME") <> "" Then
            mname = sel_pathogen_icu_rs("MID_NAME")
        End If
        If sel_pathogen_icu_rs("LAST_NAME") <> "" Then
            lname = sel_pathogen_icu_rs("LAST_NAME")
        End If
        pname = Trim(sal) & " " & Trim(FNAME) & " " & Trim(mname) & " " & Trim(lname)
        
     
        '****************** Added on 08/06/09 ********************
        If Len(Trim(sel_pathogen_icu_rs("BIRTH_DATE"))) <> 0 Then
            age = DateDiff("YYYY", sel_pathogen_icu_rs("BIRTH_DATE"), CStr(DATE11))
        End If
        '*********************************************************
            
        'If I <> 1 Then
        '    If Trim(gender) = Trim(sel_pathogen_icu_rs("SEX")) & "/" & age Then
        '        MSFlexGrid1.TextMatrix(I, 5) = ""
        '    Else
        '        MSFlexGrid1.TextMatrix(I, 5) = sel_pathogen_icu_rs("sex") & "/" & age
        '    End If
        'Else
        '    MSFlexGrid1.TextMatrix(I, 5) = sel_pathogen_icu_rs("sex") & "/" & age
        'End If
        'gender = sel_pathogen_icu_rs("sex") & "/" & age
         

        '************ Shraddha 13/06/09 ************
'        If MSFlexGrid1.TextMatrix(I, 1) <> "" Then
'            MSFlexGrid1.TextMatrix(I, 5) = Trim(sel_pathogen_icu_rs("SEX"))
'        End If
        If i <> 1 Then
            If MSFlexGrid1.TextMatrix(i, 1) = "" Then
                MSFlexGrid1.TextMatrix(i, 5) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 5) = Trim(sel_pathogen_icu_rs("SEX"))
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 5) = Trim(sel_pathogen_icu_rs("SEX"))
        End If
        '*******************************************
'        If MSFlexGrid1.TextMatrix(I, 1) <> "" Then
'            MSFlexGrid1.TextMatrix(I, 6) = age
'        End If
  
        If i <> 1 Then
            If MSFlexGrid1.TextMatrix(i, 1) = "" Then
                MSFlexGrid1.TextMatrix(i, 6) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 6) = age
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 6) = age
        End If
        '************************************************************************
        'If Trim(sel_pathogen_icu_rs("SPECIMEN")) <> "" Then
        '    If I <> 1 Then
        '        If Trim(nmaterial) = Trim(sel_pathogen_icu_rs("SPECIMEN")) Then
        '            MSFlexGrid1.TextMatrix(I, 6) = ""
        '        Else
        '            MSFlexGrid1.TextMatrix(I, 6) = sel_pathogen_icu_rs("SPECIMEN")
        '        End If
        '    Else
        '        MSFlexGrid1.TextMatrix(I, 6) = sel_pathogen_icu_rs("SPECIMEN")
        '    End If
        '    nmaterial = sel_pathogen_icu_rs("SPECIMEN")
        'End If
        
        '****************

        '************ Shraddha 13/06/09 ************
'        If MSFlexGrid1.TextMatrix(I, 1) <> "" Then
'            MSFlexGrid1.TextMatrix(I, 7) = Trim(sel_pathogen_icu_rs("SPECIMEN"))
'        End If
        
        '***** Shraddha 10/10/2011 *****
        If Trim(sel_pathogen_icu_rs("ORGSO1")) <> "" Or Trim(sel_pathogen_icu_rs("ORGSO2")) <> "" Or Trim(sel_pathogen_icu_rs("ORGSO3")) <> "" Then
            patho_grown = "  Pathogen grown : Yes"
        Else
            patho_grown = "  Pathogen grown : No"
        End If
        '*******************************
        
        If i <> 1 Then
            If Trim(UCase(SP)) = Trim(UCase(sel_pathogen_icu_rs("SPECIMEN"))) And cno = sel_pathogen_icu_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 7) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 7) = Trim(sel_pathogen_icu_rs("SPECIMEN")) & patho_grown
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 7) = Trim(sel_pathogen_icu_rs("SPECIMEN")) & patho_grown
        End If
        SP = Trim(sel_pathogen_icu_rs("SPECIMEN"))
        
        '*******************************************

        '************ Shraddha 29/07/09 ************
        'Collection Date 8
'        If Trim(sel_pathogen_icu_rs("samcoldate")) <> "" Then
'            MSFlexGrid1.TextMatrix(I, 8) = sel_pathogen_icu_rs("samcoldate")
'            colldt = sel_pathogen_icu_rs("samcoldate")
'        End If
        If i <> 1 Then
            If Trim(colldt) = Format(sel_pathogen_icu_rs("samcoldate"), "dd/mm/yyyy") And cno = sel_pathogen_icu_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 8) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 8) = Format(sel_pathogen_icu_rs("samcoldate"), "dd/mm/yyyy")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 8) = Format(sel_pathogen_icu_rs("samcoldate"), "dd/mm/yyyy")
        End If
        If (sel_pathogen_icu_rs("samcoldate")) <> Null Or (sel_pathogen_icu_rs("samcoldate")) <> "" Then
            colldt = Format(sel_pathogen_icu_rs("samcoldate"), "dd/mm/yyyy")
        End If
        '*******************************************

'        If Trim(sel_pathogen_icu_rs("ORGSO1")) <> "" Then
'            If I <> 1 Then
'                If Trim(or1) = Trim(sel_pathogen_icu_rs("ORGSO1")) Then
'                    MSFlexGrid1.TextMatrix(I, 9) = ""
'                Else
'                    MSFlexGrid1.TextMatrix(I, 9) = sel_pathogen_icu_rs("ORGSO1")
'                End If
'            Else
'                MSFlexGrid1.TextMatrix(I, 9) = sel_pathogen_icu_rs("ORGSO1")
'            End If
'            or1 = sel_pathogen_icu_rs("ORGSO1")
'        End If
                
        If Trim(sel_pathogen_icu_rs("ORGSO1")) <> "" Then
            If i <> 1 Then
                If Trim(or1) = Trim(sel_pathogen_icu_rs("ORGSO1")) And cno = sel_pathogen_icu_rs("caseno") Then
                    MSFlexGrid1.TextMatrix(i, 9) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 9) = sel_pathogen_icu_rs("ORGSO1")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 9) = sel_pathogen_icu_rs("ORGSO1")
            End If
            or1 = sel_pathogen_icu_rs("ORGSO1")
        End If
        '******************************************************
        If Trim(sel_pathogen_icu_rs("ORGSO1")) <> "" Then
            If Trim(sel_pathogen_icu_rs("ANTIBIO1")) <> "" Then
                If i <> 1 Then
                    'If Trim(antbt1) = Trim(sel_pathogen_icu_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO1")) & ")" Then
                    If Trim(antbt1) = Trim(sel_pathogen_icu_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO1")) & ")" And cno = sel_pathogen_icu_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 10) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 10) = sel_pathogen_icu_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO1")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 10) = sel_pathogen_icu_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO1")) & ")"
                End If
                antbt1 = Trim(sel_pathogen_icu_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO1")) & ")"
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 10
                If Trim(sel_pathogen_icu_rs("ANTIBIO1")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_icu_rs("ANTIBIO1")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_icu_rs("ANTIBIO1")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
                If Opt_antibiotic.Value = True Then
                    MSFlexGrid1.TextMatrix(i, 10) = Trim(sel_pathogen_icu_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO1")) & ")"
                End If
            End If
        End If
        
        '******************************************************
        If Trim(sel_pathogen_icu_rs("ORGSO2")) <> "" Then
            If i <> 1 Then
                'If Trim(or2) = Trim(sel_pathogen_icu_rs("ORGSO2")) Then
                If Trim(or2) = Trim(sel_pathogen_icu_rs("ORGSO2")) And cno = sel_pathogen_icu_rs("caseno") Then
                    MSFlexGrid1.TextMatrix(i, 11) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 11) = sel_pathogen_icu_rs("ORGSO2")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 11) = sel_pathogen_icu_rs("ORGSO2")
            End If
            or2 = sel_pathogen_icu_rs("ORGSO2")
        End If
        '******************************************************
        
        If Trim(sel_pathogen_icu_rs("ORGSO2")) <> "" Then
            If Trim(sel_pathogen_icu_rs("ANTIBIO2")) <> "" Then
                If i <> 1 Then
                    'If Trim(antbt2) = Trim(sel_pathogen_icu_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO2")) & ")" Then
                    If Trim(antbt2) = Trim(sel_pathogen_icu_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO2")) & ")" And cno = sel_pathogen_icu_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 12) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 12) = sel_pathogen_icu_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO2")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 12) = sel_pathogen_icu_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO2")) & ")"
                End If
                antbt2 = Trim(sel_pathogen_icu_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO2")) & ")"
                
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 12
                If Trim(sel_pathogen_icu_rs("ANTIBIO2")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_icu_rs("ANTIBIO2")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_icu_rs("ANTIBIO2")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
                If Opt_antibiotic.Value = True Then
                    MSFlexGrid1.TextMatrix(i, 12) = Trim(sel_pathogen_icu_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO2")) & ")"
                End If
            End If
        End If
        '******************************************************
        
        If Trim(sel_pathogen_icu_rs("ORGSO3")) <> "" Then
            If i <> 1 Then
                'If Trim(or3) = Trim(sel_pathogen_icu_rs("ORGSO3")) Then
                If Trim(or3) = Trim(sel_pathogen_icu_rs("ORGSO3")) And cno = sel_pathogen_icu_rs("caseno") Then
                    MSFlexGrid1.TextMatrix(i, 13) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 13) = sel_pathogen_icu_rs("ORGSO3")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 13) = sel_pathogen_icu_rs("ORGSO3")
            End If
            or3 = sel_pathogen_icu_rs("ORGSO3")
        End If
        
        '******************************************************
        If Trim(sel_pathogen_icu_rs("ORGSO3")) <> "" Then
            If Trim(sel_pathogen_icu_rs("ANTIBIO3")) <> "" Then
                If i <> 1 Then
                    'If Trim(antbt3) = Trim(sel_pathogen_icu_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO3")) & ")" Then
                    If Trim(antbt3) = Trim(sel_pathogen_icu_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO3")) & ")" And cno = sel_pathogen_icu_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 14) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 14) = sel_pathogen_icu_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO3")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 14) = sel_pathogen_icu_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO3")) & ")"
                End If
                antbt3 = Trim(sel_pathogen_icu_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO3")) & ")"
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 14
                If Trim(sel_pathogen_icu_rs("ANTIBIO3")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_icu_rs("ANTIBIO3")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_icu_rs("ANTIBIO3")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
                If Opt_antibiotic.Value = True Then
                    MSFlexGrid1.TextMatrix(i, 14) = Trim(sel_pathogen_icu_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_icu_rs("ANTIBIO3")) & ")"
                End If
            End If
        End If
        '******************************************************
        cno = sel_pathogen_icu_rs("caseno")
        sel_pathogen_icu_rs.MoveNext
    Next
 ''''''''''''''''''''''''''''''''''''''''''''''''''''
End If

'***************** Patient option ***********************
If Option3.Value = True Then
    c_no = MaskEdBox3.FormattedText
    sel_pathogen_patient_qr(0) = c_no
    sel_pathogen_patient_qr(1) = var_date_f
    sel_pathogen_patient_qr(2) = var_date_t
    Set sel_pathogen_patient_rs = sel_pathogen_patient_qr.OpenResultset(1, 3)
    
    '***************************************
    MSFlexGrid1.Cols = 16

    MSFlexGrid1.TextMatrix(0, 0) = "SNO."
    MSFlexGrid1.TextMatrix(0, 1) = "CASE NO."
    MSFlexGrid1.TextMatrix(0, 2) = "LAB NO."
    MSFlexGrid1.TextMatrix(0, 3) = "REQUISITION NO."
    MSFlexGrid1.TextMatrix(0, 4) = "WARD/ICU"
    MSFlexGrid1.TextMatrix(0, 5) = "NAME"
    MSFlexGrid1.TextMatrix(0, 6) = "SEX"
    MSFlexGrid1.TextMatrix(0, 7) = "AGE" '29/07/09
    MSFlexGrid1.TextMatrix(0, 8) = "SPECIMEN"
    MSFlexGrid1.TextMatrix(0, 9) = "COLLECTION DATE" 'Collectoin Date
    MSFlexGrid1.TextMatrix(0, 10) = "ORGANISM1"
    MSFlexGrid1.TextMatrix(0, 11) = "ANTIBIOTIC1"
    MSFlexGrid1.TextMatrix(0, 12) = "ORGANISM2"
    MSFlexGrid1.TextMatrix(0, 13) = "ANTIBIOTIC2"
    MSFlexGrid1.TextMatrix(0, 14) = "ORGANISM3"
    MSFlexGrid1.TextMatrix(0, 15) = "ANTIBIOTIC3"
    MSFlexGrid1.RowHeight(0) = 350
    '***************************************
    
    
    i = sel_pathogen_patient_rs.RowCount
    If i = 0 Then
        title_chk = True
    Else
        title_chk = False
    End If
    
    MSFlexGrid1.Rows = sel_pathogen_patient_rs.RowCount + 1
    For i = 1 To sel_pathogen_patient_rs.RowCount
        MSFlexGrid1.ColWidth(0) = 500
        MSFlexGrid1.ColWidth(1) = 1000
        MSFlexGrid1.ColWidth(2) = 900 'Labno
        MSFlexGrid1.ColWidth(3) = 1500
        MSFlexGrid1.ColWidth(4) = 2400
        MSFlexGrid1.ColWidth(5) = 2000 'ICU/Ward
        MSFlexGrid1.ColWidth(6) = 500 'Gender
        MSFlexGrid1.ColWidth(7) = 500 'age
        MSFlexGrid1.ColWidth(8) = 2000
        MSFlexGrid1.ColWidth(9) = 2000 'Collection Date
        MSFlexGrid1.ColWidth(10) = 2400
        MSFlexGrid1.ColWidth(11) = 2400
        MSFlexGrid1.ColWidth(12) = 2400
        MSFlexGrid1.ColWidth(13) = 2400
        MSFlexGrid1.ColWidth(14) = 2400
        MSFlexGrid1.ColWidth(15) = 1900
        MSFlexGrid1.RowHeight(i) = 350
        '''''''''''''''''''''''''''''''''''''''

        If i <> 1 Then
            If Trim(cno) = Trim(sel_pathogen_patient_rs("caseno")) Then
                MSFlexGrid1.TextMatrix(i, 1) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 1) = sel_pathogen_patient_rs("caseno")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 1) = sel_pathogen_patient_rs("caseno")
        End If
        'cno = sel_pathogen_patient_rs("caseno")


        If MSFlexGrid1.TextMatrix(i, 1) <> "" Then
            MSFlexGrid1.TextMatrix(i, 0) = n
            n = n + 1
        End If
        
        '*********** Added by Shraddha on 12/06/09 ***********
        sel_labno_qr(0) = Trim(sel_pathogen_patient_rs("reqnno"))
        Set sel_labno_rs = sel_labno_qr.OpenResultset(1, 3)
         
        If i <> 1 Then
            'If Trim(lab_no) = Trim(sel_pathogen_patient_rs("pathno")) Then
            If Trim(lab_no) = Trim(sel_pathogen_patient_rs("pathno")) And cno = sel_pathogen_patient_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 2) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 2) = Trim(sel_pathogen_patient_rs("pathno"))
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 2) = Trim(sel_pathogen_patient_rs("pathno"))
        End If
        lab_no = Trim(sel_pathogen_patient_rs("pathno"))
        '*****************************************************
    
        If i <> 1 Then
            'If Trim(rno) = Trim(sel_pathogen_patient_rs("reqnno")) Then
            If Trim(rno) = Trim(sel_pathogen_patient_rs("reqnno")) And cno = sel_pathogen_patient_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 3) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 3) = sel_pathogen_patient_rs("reqnno")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 3) = sel_pathogen_patient_rs("reqnno")
        End If
        rno = sel_pathogen_patient_rs("reqnno")
                  
        '***************************************************
        If MSFlexGrid1.TextMatrix(i, 3) <> "" Then
            If sel_pathogen_patient_rs("ICDS2") = "YES" Then
                MSFlexGrid1.TextMatrix(i, 4) = "ICU"
            End If
            If sel_pathogen_patient_rs("WARD_NAME") <> "" Then
                MSFlexGrid1.TextMatrix(i, 4) = sel_pathogen_patient_rs("WARD_NAME")
            End If
        End If
        '****************************************************

        If i <> 1 Then
            'If Trim(pname) = Trim(Trim(sel_pathogen_patient_rs("tag")) & " " & Trim(sel_pathogen_patient_rs("FIRST_NAME")) & " " & Trim(sel_pathogen_patient_rs("MID_NAME")) & " " & Trim(sel_pathogen_patient_rs("LAST_NAME"))) Then
            If cno = sel_pathogen_patient_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 5) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 5) = Trim(sel_pathogen_patient_rs("tag")) & " " & Trim(sel_pathogen_patient_rs("FIRST_NAME")) & " " & Trim(sel_pathogen_patient_rs("MID_NAME")) & " " & Trim(sel_pathogen_patient_rs("LAST_NAME"))
            End If
        Else
            MSFlexGrid1.TextMatrix(i, 5) = Trim(sel_pathogen_patient_rs("tag")) & " " & Trim(sel_pathogen_patient_rs("FIRST_NAME")) & " " & Trim(sel_pathogen_patient_rs("MID_NAME")) & " " & Trim(sel_pathogen_patient_rs("LAST_NAME"))
        End If

        If sel_pathogen_patient_rs("tag") <> "" Then
            sal = sel_pathogen_patient_rs("tag")
        End If
        If sel_pathogen_patient_rs("FIRST_NAME") <> "" Then
            FNAME = sel_pathogen_patient_rs("FIRST_NAME")
        End If
        If sel_pathogen_patient_rs("MID_NAME") <> "" Then
            mname = sel_pathogen_patient_rs("MID_NAME")
        End If
        If sel_pathogen_patient_rs("LAST_NAME") <> "" Then
            lname = sel_pathogen_patient_rs("LAST_NAME")
        End If
        pname = Trim(sal) & " " & Trim(FNAME) & " " & Trim(mname) & " " & Trim(lname)
     
  
        '****************** Added on 08/06/09 ********************
        If Len(Trim(sel_pathogen_patient_rs("BIRTH_DATE"))) <> 0 Then
            age = DateDiff("YYYY", sel_pathogen_patient_rs("BIRTH_DATE"), CStr(DATE11))
        End If
        
        '*********************************************************
          
        'If I <> 1 Then
        '    If Trim(gender) = Trim(sel_pathogen_patient_rs("SEX")) & "/" & age Then
        '        MSFlexGrid1.TextMatrix(I, 6) = ""
        '    Else
        '        MSFlexGrid1.TextMatrix(I, 6) = sel_pathogen_patient_rs("sex") & "/" & age
        '    End If
        'Else
        '        MSFlexGrid1.TextMatrix(I, 6) = sel_pathogen_patient_rs("sex") & "/" & age
        'End If
        'gender = sel_pathogen_patient_rs("sex") & "/" & age
        
        '************ Shraddha 13/06/09 ************
        If MSFlexGrid1.TextMatrix(i, 1) <> "" Then
            'MSFlexGrid1.TextMatrix(I, 6) = Trim(sel_pathogen_patient_rs("SEX")) & "/" & age
            MSFlexGrid1.TextMatrix(i, 6) = Trim(sel_pathogen_patient_rs("SEX"))
        End If
        
        
        
        
        If MSFlexGrid1.TextMatrix(i, 1) <> "" Then
            MSFlexGrid1.TextMatrix(i, 7) = age
        End If
        '*******************************************
     
        'If I <> 1 Then
        '    If Trim(nmaterial) = Trim(sel_pathogen_patient_rs("SPECIMEN")) Then
        '        MSFlexGrid1.TextMatrix(I, 7) = ""
        '    Else
        '        MSFlexGrid1.TextMatrix(I, 7) = sel_pathogen_patient_rs("SPECIMEN")
        '    End If
        'Else
        '    MSFlexGrid1.TextMatrix(I, 7) = sel_pathogen_patient_rs("SPECIMEN")
        'End If
        'nmaterial = sel_pathogen_patient_rs("SPECIMEN")
        
        '************ Shraddha 13/06/09 ************
'        If MSFlexGrid1.TextMatrix(I, 1) <> "" Then
'            MSFlexGrid1.TextMatrix(I, 8) = Trim(sel_pathogen_patient_rs("SPECIMEN"))
'        End If

         '***** Shraddha 10/10/2011 *****
        If Trim(sel_pathogen_patient_rs("ORGSO1")) <> "" Or Trim(sel_pathogen_patient_rs("ORGSO2")) <> "" Or Trim(sel_pathogen_patient_rs("ORGSO3")) <> "" Then
            patho_grown = "  Pathogen grown : Yes"
        Else
            patho_grown = "  Pathogen grown : No"
        End If
        '*******************************
        
        If i <> 1 Then
            If Trim(UCase(SP)) = Trim(UCase(sel_pathogen_patient_rs("SPECIMEN"))) And cno = sel_pathogen_patient_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 8) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 8) = Trim(sel_pathogen_patient_rs("SPECIMEN")) & patho_grown
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 8) = Trim(sel_pathogen_patient_rs("SPECIMEN")) & patho_grown
        End If
        SP = Trim(sel_pathogen_patient_rs("SPECIMEN"))
        '*******************************************
     
        '*********** Shraddha Pandit 29/07/09 ************
        'Collection Date 9
'        If Trim(sel_pathogen_patient_rs("samcoldate")) <> "" Then
'            MSFlexGrid1.TextMatrix(I, 9) = sel_pathogen_patient_rs("samcoldate")
'            colldt = sel_pathogen_patient_rs("samcoldate")
'        End If
        If i <> 1 Then
            If Trim(colldt) = Format(sel_pathogen_patient_rs("samcoldate"), "dd/mm/yyyy") And cno = sel_pathogen_patient_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 9) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 9) = Format(sel_pathogen_patient_rs("samcoldate"), "dd/mm/yyyy")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 9) = Format(sel_pathogen_patient_rs("samcoldate"), "dd/mm/yyyy")
        End If
        If (sel_pathogen_patient_rs("samcoldate")) <> Null Or (sel_pathogen_patient_rs("samcoldate")) <> "" Then
            colldt = Format(sel_pathogen_patient_rs("samcoldate"), "dd/mm/yyyy")
        End If
        '*************************************************
'        If Trim(sel_pathogen_patient_rs("ORGSO1")) <> "" Then
'            If I <> 1 Then
'                If Trim(or1) = Trim(sel_pathogen_patient_rs("ORGSO1")) Then
'                    MSFlexGrid1.TextMatrix(I, 10) = ""
'                Else
'                    MSFlexGrid1.TextMatrix(I, 10) = sel_pathogen_patient_rs("ORGSO1")
'                End If
'            Else
'                MSFlexGrid1.TextMatrix(I, 10) = sel_pathogen_patient_rs("ORGSO1")
'            End If
'            or1 = sel_pathogen_patient_rs("ORGSO1")
'        End If
        
        If Trim(sel_pathogen_patient_rs("ORGSO1")) <> "" Then
            If i <> 1 Then
                'If Trim(or1) = Trim(sel_pathogen_ward_rs("ORGSO1")) Then
                If Trim(or1) = Trim(sel_pathogen_patient_rs("ORGSO1")) And cno = sel_pathogen_patient_rs("caseno") Then
                
                'If MSFlexGrid1.TextMatrix(I, 1) = "" Then
                    MSFlexGrid1.TextMatrix(i, 10) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 10) = sel_pathogen_patient_rs("ORGSO1")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 10) = sel_pathogen_patient_rs("ORGSO1")
            End If
            or1 = sel_pathogen_patient_rs("ORGSO1")
        End If

        If Trim(sel_pathogen_patient_rs("ORGSO1")) <> "" Then
            If Trim(sel_pathogen_patient_rs("ANTIBIO1")) <> "" Then
                If i <> 1 Then
                    'If Trim(antbt1) = Trim(sel_pathogen_patient_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO1")) & ")" Then
                    If Trim(antbt1) = Trim(sel_pathogen_patient_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO1")) & ")" And cno = sel_pathogen_patient_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 11) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 11) = sel_pathogen_patient_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO1")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 11) = sel_pathogen_patient_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO1")) & ")"
                End If
                antbt1 = Trim(sel_pathogen_patient_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO1")) & ")"
                
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 11
                If Trim(sel_pathogen_patient_rs("ANTIBIO1")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_patient_rs("ANTIBIO1")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_patient_rs("ANTIBIO1")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
            End If
        End If

        
        If Trim(sel_pathogen_patient_rs("ORGSO2")) <> "" Then
            If i <> 1 Then
                'If Trim(or2) = Trim(sel_pathogen_patient_rs("ORGSO2")) Then
                If Trim(or2) = Trim(sel_pathogen_patient_rs("ORGSO2")) And cno = sel_pathogen_patient_rs("caseno") Then
                    MSFlexGrid1.TextMatrix(i, 12) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 12) = sel_pathogen_patient_rs("ORGSO2")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 12) = sel_pathogen_patient_rs("ORGSO2")
            End If
            or2 = sel_pathogen_patient_rs("ORGSO2")
        End If
        


        If Trim(sel_pathogen_patient_rs("ORGSO2")) <> "" Then
            If Trim(sel_pathogen_patient_rs("ANTIBIO2")) <> "" Then
                If i <> 1 Then
                    'If Trim(antbt2) = Trim(sel_pathogen_patient_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO2")) & ")" Then
                    If Trim(antbt2) = Trim(sel_pathogen_patient_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO2")) & ")" And cno = sel_pathogen_patient_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 13) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 13) = sel_pathogen_patient_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO2")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 13) = sel_pathogen_patient_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO2")) & ")"
                End If
                antbt2 = Trim(sel_pathogen_patient_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO2")) & ")"
                
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 13
                If Trim(sel_pathogen_patient_rs("ANTIBIO2")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_patient_rs("ANTIBIO2")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_patient_rs("ANTIBIO2")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
            End If
        End If

        
        If Trim(sel_pathogen_patient_rs("ORGSO3")) <> "" Then
            If i <> 1 Then
                'If Trim(or3) = Trim(sel_pathogen_patient_rs("ORGSO3")) Then
                If Trim(or3) = Trim(sel_pathogen_patient_rs("ORGSO3")) And cno = sel_pathogen_patient_rs("caseno") Then
                    MSFlexGrid1.TextMatrix(i, 14) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 14) = sel_pathogen_patient_rs("ORGSO3")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 14) = sel_pathogen_patient_rs("ORGSO3")
            End If
            or3 = sel_pathogen_patient_rs("ORGSO3")
        End If


        If Trim(sel_pathogen_patient_rs("ORGSO3")) <> "" Then
            If Trim(sel_pathogen_patient_rs("ANTIBIO3")) <> "" Then
                If i <> 1 Then
                    'If Trim(antbt3) = Trim(sel_pathogen_patient_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO3")) & ")" Then
                    If Trim(antbt3) = Trim(sel_pathogen_patient_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO3")) & ")" And cno = sel_pathogen_patient_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 15) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 15) = sel_pathogen_patient_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO3")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 15) = sel_pathogen_patient_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO3")) & ")"
                End If
                antbt3 = Trim(sel_pathogen_patient_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_patient_rs("ANTIBIO3")) & ")"
                
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 15
                If Trim(sel_pathogen_patient_rs("ANTIBIO3")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_patient_rs("ANTIBIO3")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_patient_rs("ANTIBIO3")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
            End If
        End If

        cno = sel_pathogen_patient_rs("caseno")
        sel_pathogen_patient_rs.MoveNext
    Next

End If
'********************************************************

'*********** Shraddha Pandit 31/07/09 ****************
'***************** Unit option ***********************
If Option4.Value = True Then

    If cmb_doctor.Visible = True Then 'Doctor
        unit = Trim(cmb_unit.Text)
        doctor = Trim(cmb_doctor.Text)
        sql_pathogen_unit = "select * from RISLIB.pathogen_new where REFBYDOC=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
        Set sel_pathogen_unit_qr = rdb.CreateQuery("", sql_pathogen_unit)
        sel_pathogen_unit_qr(0) = doctor
        sel_pathogen_unit_qr(1) = var_date_f
        sel_pathogen_unit_qr(2) = var_date_t
        Set sel_pathogen_unit_rs = sel_pathogen_unit_qr.OpenResultset(1, 3)
        
    Else
        unit = Trim(cmb_unit.Text)
        If unit = "ALL" Then
            sql_pathogen_unit = "select * from RISLIB.pathogen_new where REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
            Set sel_pathogen_unit_qr = rdb.CreateQuery("", sql_pathogen_unit)
            
            sel_pathogen_unit_qr(0) = var_date_f
            sel_pathogen_unit_qr(1) = var_date_t
            Set sel_pathogen_unit_rs = sel_pathogen_unit_qr.OpenResultset(1, 3)
        Else
            sql_pathogen_unit = "select * from RISLIB.pathogen_new where REFBYDOC=? AND REQNDT>=? AND REQNDT<=? ORDER BY CASENO, REQNNO, FIRST_NAME, MID_NAME"
            Set sel_pathogen_unit_qr = rdb.CreateQuery("", sql_pathogen_unit)
            sel_pathogen_unit_qr(0) = unit
            sel_pathogen_unit_qr(1) = var_date_f
            sel_pathogen_unit_qr(2) = var_date_t
            Set sel_pathogen_unit_rs = sel_pathogen_unit_qr.OpenResultset(1, 3)
        End If
    End If
    
    MSFlexGrid1.Cols = 16

    MSFlexGrid1.TextMatrix(0, 0) = "SNO."
    MSFlexGrid1.TextMatrix(0, 1) = "CASE NO."
    MSFlexGrid1.TextMatrix(0, 2) = "LAB NO."
    MSFlexGrid1.TextMatrix(0, 3) = "REQUISITION NO."
    MSFlexGrid1.TextMatrix(0, 4) = "WARD/ICU"
    MSFlexGrid1.TextMatrix(0, 5) = "NAME"
    MSFlexGrid1.TextMatrix(0, 6) = "SEX"
    MSFlexGrid1.TextMatrix(0, 7) = "AGE" '29/07/09
    MSFlexGrid1.TextMatrix(0, 8) = "SPECIMEN"
    MSFlexGrid1.TextMatrix(0, 9) = "COLLECTION DATE" 'Collectoin Date
    MSFlexGrid1.TextMatrix(0, 10) = "ORGANISM1"
    MSFlexGrid1.TextMatrix(0, 11) = "ANTIBIOTIC1"
    MSFlexGrid1.TextMatrix(0, 12) = "ORGANISM2"
    MSFlexGrid1.TextMatrix(0, 13) = "ANTIBIOTIC2"
    MSFlexGrid1.TextMatrix(0, 14) = "ORGANISM3"
    MSFlexGrid1.TextMatrix(0, 15) = "ANTIBIOTIC3"
    MSFlexGrid1.RowHeight(0) = 350
    '***************************************
    
    i = sel_pathogen_unit_rs.RowCount
    If i = 0 Then
        title_chk = True
    Else
        title_chk = False
    End If
    
    MSFlexGrid1.Rows = sel_pathogen_unit_rs.RowCount + 1
    For i = 1 To sel_pathogen_unit_rs.RowCount
        MSFlexGrid1.ColWidth(0) = 500
        MSFlexGrid1.ColWidth(1) = 1000
        MSFlexGrid1.ColWidth(2) = 900 'Labno
        MSFlexGrid1.ColWidth(3) = 1500
        MSFlexGrid1.ColWidth(4) = 2400
        MSFlexGrid1.ColWidth(5) = 2000 'ICU/Ward
        MSFlexGrid1.ColWidth(6) = 500 'Gender
        MSFlexGrid1.ColWidth(7) = 500 'age
        MSFlexGrid1.ColWidth(8) = 2000
        MSFlexGrid1.ColWidth(9) = 2000 'Collection Date
        MSFlexGrid1.ColWidth(10) = 2400
        MSFlexGrid1.ColWidth(11) = 2400
        MSFlexGrid1.ColWidth(12) = 2400
        MSFlexGrid1.ColWidth(13) = 2400
        MSFlexGrid1.ColWidth(14) = 2400
        MSFlexGrid1.ColWidth(15) = 1900
        MSFlexGrid1.RowHeight(i) = 350
        '''''''''''''''''''''''''''''''''''''''

        If i <> 1 Then
            If Trim(cno) = Trim(sel_pathogen_unit_rs("caseno")) Then
                MSFlexGrid1.TextMatrix(i, 1) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 1) = sel_pathogen_unit_rs("caseno")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 1) = sel_pathogen_unit_rs("caseno")
        End If
        
        If MSFlexGrid1.TextMatrix(i, 1) <> "" Then
            MSFlexGrid1.TextMatrix(i, 0) = n
            n = n + 1
        End If
                
        sel_labno_qr(0) = Trim(sel_pathogen_unit_rs("reqnno"))
        Set sel_labno_rs = sel_labno_qr.OpenResultset(1, 3)
         
        If i <> 1 Then
            If Trim(lab_no) = Trim(sel_pathogen_unit_rs("pathno")) And cno = sel_pathogen_unit_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 2) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 2) = Trim(sel_pathogen_unit_rs("pathno"))
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 2) = Trim(sel_pathogen_unit_rs("pathno"))
        End If
        lab_no = Trim(sel_pathogen_unit_rs("pathno"))
            
        If i <> 1 Then
            If Trim(rno) = Trim(sel_pathogen_unit_rs("reqnno")) And cno = sel_pathogen_unit_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 3) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 3) = sel_pathogen_unit_rs("reqnno")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 3) = sel_pathogen_unit_rs("reqnno")
        End If
        rno = sel_pathogen_unit_rs("reqnno")
                
        If MSFlexGrid1.TextMatrix(i, 3) <> "" Then
            If sel_pathogen_unit_rs("ICDS2") = "YES" Then
                MSFlexGrid1.TextMatrix(i, 4) = "ICU"
            End If
            If sel_pathogen_unit_rs("WARD_NAME") <> "" Then
                MSFlexGrid1.TextMatrix(i, 4) = sel_pathogen_unit_rs("WARD_NAME")
            End If
        End If
        
        If i <> 1 Then
            If cno = sel_pathogen_unit_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 5) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 5) = Trim(sel_pathogen_unit_rs("tag")) & " " & Trim(sel_pathogen_unit_rs("FIRST_NAME")) & " " & Trim(sel_pathogen_unit_rs("MID_NAME")) & " " & Trim(sel_pathogen_unit_rs("LAST_NAME"))
            End If
        Else
            MSFlexGrid1.TextMatrix(i, 5) = Trim(sel_pathogen_unit_rs("tag")) & " " & Trim(sel_pathogen_unit_rs("FIRST_NAME")) & " " & Trim(sel_pathogen_unit_rs("MID_NAME")) & " " & Trim(sel_pathogen_unit_rs("LAST_NAME"))
        End If

        If sel_pathogen_unit_rs("tag") <> "" Then
            sal = sel_pathogen_unit_rs("tag")
        End If
        If sel_pathogen_unit_rs("FIRST_NAME") <> "" Then
            FNAME = sel_pathogen_unit_rs("FIRST_NAME")
        End If
        If sel_pathogen_unit_rs("MID_NAME") <> "" Then
            mname = sel_pathogen_unit_rs("MID_NAME")
        End If
        If sel_pathogen_unit_rs("LAST_NAME") <> "" Then
            lname = sel_pathogen_unit_rs("LAST_NAME")
        End If
        pname = Trim(sal) & " " & Trim(FNAME) & " " & Trim(mname) & " " & Trim(lname)
             
        If Len(Trim(sel_pathogen_unit_rs("BIRTH_DATE"))) <> 0 Then
            age = DateDiff("YYYY", sel_pathogen_unit_rs("BIRTH_DATE"), CStr(DATE11))
        End If
          
        If MSFlexGrid1.TextMatrix(i, 1) <> "" Then
            MSFlexGrid1.TextMatrix(i, 6) = Trim(sel_pathogen_unit_rs("SEX"))
        End If
        
        If MSFlexGrid1.TextMatrix(i, 1) <> "" Then
            MSFlexGrid1.TextMatrix(i, 7) = age
        End If
        
         '***** Shraddha 10/10/2011 *****
        If Trim(sel_pathogen_unit_rs("ORGSO1")) <> "" Or Trim(sel_pathogen_unit_rs("ORGSO2")) <> "" Or Trim(sel_pathogen_unit_rs("ORGSO3")) <> "" Then
            patho_grown = "  Pathogen grown : Yes"
        Else
            patho_grown = "  Pathogen grown : No"
        End If
        '*******************************
                
        If i <> 1 Then
            If Trim(UCase(SP)) = Trim(UCase(sel_pathogen_unit_rs("SPECIMEN"))) And cno = sel_pathogen_unit_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 8) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 8) = Trim(sel_pathogen_unit_rs("SPECIMEN")) & patho_grown
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 8) = Trim(sel_pathogen_unit_rs("SPECIMEN")) & patho_grown
        End If
        SP = Trim(sel_pathogen_unit_rs("SPECIMEN"))
        
        If i <> 1 Then
            If Trim(colldt) = Format(sel_pathogen_unit_rs("samcoldate"), "dd/mm/yyyy") And cno = sel_pathogen_unit_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 9) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 9) = Format(sel_pathogen_unit_rs("samcoldate"), "dd/mm/yyyy")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 9) = Format(sel_pathogen_unit_rs("samcoldate"), "dd/mm/yyyy")
        End If
        If (sel_pathogen_unit_rs("samcoldate")) <> Null Or (sel_pathogen_unit_rs("samcoldate")) <> "" Then
            colldt = Format(sel_pathogen_unit_rs("samcoldate"), "dd/mm/yyyy")
        End If
                
        If Trim(sel_pathogen_unit_rs("ORGSO1")) <> "" Then
            If i <> 1 Then
                If Trim(or1) = Trim(sel_pathogen_unit_rs("ORGSO1")) And cno = sel_pathogen_unit_rs("caseno") Then
                    MSFlexGrid1.TextMatrix(i, 10) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 10) = sel_pathogen_unit_rs("ORGSO1")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 10) = sel_pathogen_unit_rs("ORGSO1")
            End If
            or1 = sel_pathogen_unit_rs("ORGSO1")
        End If

        If Trim(sel_pathogen_unit_rs("ORGSO1")) <> "" Then
            If Trim(sel_pathogen_unit_rs("ANTIBIO1")) <> "" Then
                If i <> 1 Then
                    If Trim(antbt1) = Trim(sel_pathogen_unit_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO1")) & ")" And cno = sel_pathogen_unit_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 11) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 11) = sel_pathogen_unit_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO1")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 11) = sel_pathogen_unit_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO1")) & ")"
                End If
                antbt1 = Trim(sel_pathogen_unit_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO1")) & ")"
                
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 11
                If Trim(sel_pathogen_unit_rs("ANTIBIO1")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_unit_rs("ANTIBIO1")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_unit_rs("ANTIBIO1")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
            End If
        End If
        
        If Trim(sel_pathogen_unit_rs("ORGSO2")) <> "" Then
            If i <> 1 Then
                If Trim(or2) = Trim(sel_pathogen_unit_rs("ORGSO2")) And cno = sel_pathogen_unit_rs("caseno") Then
                    MSFlexGrid1.TextMatrix(i, 12) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 12) = sel_pathogen_unit_rs("ORGSO2")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 12) = sel_pathogen_unit_rs("ORGSO2")
            End If
            or2 = sel_pathogen_unit_rs("ORGSO2")
        End If
        
        If Trim(sel_pathogen_unit_rs("ORGSO2")) <> "" Then
            If Trim(sel_pathogen_unit_rs("ANTIBIO2")) <> "" Then
                If i <> 1 Then
                    If Trim(antbt2) = Trim(sel_pathogen_unit_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO2")) & ")" And cno = sel_pathogen_unit_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 13) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 13) = sel_pathogen_unit_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO2")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 13) = sel_pathogen_unit_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO2")) & ")"
                End If
                antbt2 = Trim(sel_pathogen_unit_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO2")) & ")"
                
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 13
                If Trim(sel_pathogen_unit_rs("ANTIBIO2")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_unit_rs("ANTIBIO2")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_unit_rs("ANTIBIO2")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
            End If
        End If
        
        If Trim(sel_pathogen_unit_rs("ORGSO3")) <> "" Then
            If i <> 1 Then
                If Trim(or3) = Trim(sel_pathogen_unit_rs("ORGSO3")) And cno = sel_pathogen_unit_rs("caseno") Then
                    MSFlexGrid1.TextMatrix(i, 14) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 14) = sel_pathogen_unit_rs("ORGSO3")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 14) = sel_pathogen_unit_rs("ORGSO3")
            End If
            or3 = sel_pathogen_unit_rs("ORGSO3")
        End If

        If Trim(sel_pathogen_unit_rs("ORGSO3")) <> "" Then
            If Trim(sel_pathogen_unit_rs("ANTIBIO3")) <> "" Then
                If i <> 1 Then
                    If Trim(antbt3) = Trim(sel_pathogen_unit_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO3")) & ")" And cno = sel_pathogen_unit_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 15) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 15) = sel_pathogen_unit_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO3")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 15) = sel_pathogen_unit_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO3")) & ")"
                End If
                antbt3 = Trim(sel_pathogen_unit_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO3")) & ")"
                
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 15
                If Trim(sel_pathogen_unit_rs("ANTIBIO3")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_unit_rs("ANTIBIO3")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_unit_rs("ANTIBIO3")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
            End If
        End If

        cno = sel_pathogen_unit_rs("caseno")
        sel_pathogen_unit_rs.MoveNext
    Next
    
    sel_pathogen_unit_rs.Close
    sel_pathogen_unit_qr.Close

End If

'********************************************************



Call opproc


'SFlexGrid1.Rows = 2
'If MSFlexGrid1.TextMatrix(1, 1) <> "" Then
    Frame_colour.Visible = True
'End If
lbl_caption.Visible = True

Call proc1


End Sub
Public Sub opproc()
'***** Shraddha 08/11/2011 *****
If opt_OP.Value = True Then

sql_pathogen_unit = "select * from RISLIB.pathogen_new where PATNSTAT='OP' AND REQNDT>=? AND REQNDT<=? ORDER BY PATHNO " 'CASENO, REQNNO, FIRST_NAME, MID_NAME"
Set sel_pathogen_unit_qr = rdb.CreateQuery("", sql_pathogen_unit)

sel_pathogen_unit_qr(0) = var_date_f
sel_pathogen_unit_qr(1) = var_date_t
Set sel_pathogen_unit_rs = sel_pathogen_unit_qr.OpenResultset(1, 3)

MSFlexGrid1.Cols = 16

    MSFlexGrid1.TextMatrix(0, 0) = "SNO."
    MSFlexGrid1.TextMatrix(0, 1) = "CASE NO."
    MSFlexGrid1.TextMatrix(0, 2) = "LAB NO."
    MSFlexGrid1.TextMatrix(0, 3) = "REQUISITION NO."
    MSFlexGrid1.TextMatrix(0, 4) = "WARD/ICU"
    MSFlexGrid1.TextMatrix(0, 5) = "NAME"
    MSFlexGrid1.TextMatrix(0, 6) = "SEX"
    MSFlexGrid1.TextMatrix(0, 7) = "AGE" '29/07/09
    MSFlexGrid1.TextMatrix(0, 8) = "SPECIMEN"
    MSFlexGrid1.TextMatrix(0, 9) = "COLLECTION DATE" 'Collectoin Date
    MSFlexGrid1.TextMatrix(0, 10) = "ORGANISM1"
    MSFlexGrid1.TextMatrix(0, 11) = "ANTIBIOTIC1"
    MSFlexGrid1.TextMatrix(0, 12) = "ORGANISM2"
    MSFlexGrid1.TextMatrix(0, 13) = "ANTIBIOTIC2"
    MSFlexGrid1.TextMatrix(0, 14) = "ORGANISM3"
    MSFlexGrid1.TextMatrix(0, 15) = "ANTIBIOTIC3"
    MSFlexGrid1.RowHeight(0) = 350
    '***************************************
    
    i = sel_pathogen_unit_rs.RowCount
    If i = 0 Then
        title_chk = True
    Else
        title_chk = False
    End If
    
    MSFlexGrid1.Rows = sel_pathogen_unit_rs.RowCount + 1
    For i = 1 To sel_pathogen_unit_rs.RowCount
        MSFlexGrid1.ColWidth(0) = 500
        MSFlexGrid1.ColWidth(1) = 1000
        MSFlexGrid1.ColWidth(2) = 900 'Labno
        MSFlexGrid1.ColWidth(3) = 1500
        MSFlexGrid1.ColWidth(4) = 2400
        MSFlexGrid1.ColWidth(5) = 2000 'ICU/Ward
        MSFlexGrid1.ColWidth(6) = 500 'Gender
        MSFlexGrid1.ColWidth(7) = 500 'age
        MSFlexGrid1.ColWidth(8) = 2000
        MSFlexGrid1.ColWidth(9) = 2000 'Collection Date
        MSFlexGrid1.ColWidth(10) = 2400
        MSFlexGrid1.ColWidth(11) = 2400
        MSFlexGrid1.ColWidth(12) = 2400
        MSFlexGrid1.ColWidth(13) = 2400
        MSFlexGrid1.ColWidth(14) = 2400
        MSFlexGrid1.ColWidth(15) = 1900
        MSFlexGrid1.RowHeight(i) = 350
        '''''''''''''''''''''''''''''''''''''''

        If i <> 1 Then
            If Trim(cno) = Trim(sel_pathogen_unit_rs("caseno")) Then
                MSFlexGrid1.TextMatrix(i, 1) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 1) = sel_pathogen_unit_rs("caseno")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 1) = sel_pathogen_unit_rs("caseno")
        End If
        
        If MSFlexGrid1.TextMatrix(i, 1) <> "" Then
            MSFlexGrid1.TextMatrix(i, 0) = n
            n = n + 1
        End If
                
        sel_labno_qr(0) = Trim(sel_pathogen_unit_rs("reqnno"))
        Set sel_labno_rs = sel_labno_qr.OpenResultset(1, 3)
         
        If i <> 1 Then
            If Trim(lab_no) = Trim(sel_pathogen_unit_rs("pathno")) And cno = sel_pathogen_unit_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 2) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 2) = Trim(sel_pathogen_unit_rs("pathno"))
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 2) = Trim(sel_pathogen_unit_rs("pathno"))
        End If
        lab_no = Trim(sel_pathogen_unit_rs("pathno"))
            
        If i <> 1 Then
            If Trim(rno) = Trim(sel_pathogen_unit_rs("reqnno")) And cno = sel_pathogen_unit_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 3) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 3) = sel_pathogen_unit_rs("reqnno")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 3) = sel_pathogen_unit_rs("reqnno")
        End If
        rno = sel_pathogen_unit_rs("reqnno")
                
        If MSFlexGrid1.TextMatrix(i, 3) <> "" Then
            If sel_pathogen_unit_rs("ICDS2") = "YES" Then
                MSFlexGrid1.TextMatrix(i, 4) = "ICU"
            End If
            If sel_pathogen_unit_rs("WARD_NAME") <> "" Then
                MSFlexGrid1.TextMatrix(i, 4) = sel_pathogen_unit_rs("WARD_NAME")
            End If
        End If
        
        If i <> 1 Then
            If cno = sel_pathogen_unit_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 5) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 5) = Trim(sel_pathogen_unit_rs("tag")) & " " & Trim(sel_pathogen_unit_rs("FIRST_NAME")) & " " & Trim(sel_pathogen_unit_rs("MID_NAME")) & " " & Trim(sel_pathogen_unit_rs("LAST_NAME"))
            End If
        Else
            MSFlexGrid1.TextMatrix(i, 5) = Trim(sel_pathogen_unit_rs("tag")) & " " & Trim(sel_pathogen_unit_rs("FIRST_NAME")) & " " & Trim(sel_pathogen_unit_rs("MID_NAME")) & " " & Trim(sel_pathogen_unit_rs("LAST_NAME"))
        End If

        If sel_pathogen_unit_rs("tag") <> "" Then
            sal = sel_pathogen_unit_rs("tag")
        End If
        If sel_pathogen_unit_rs("FIRST_NAME") <> "" Then
            FNAME = sel_pathogen_unit_rs("FIRST_NAME")
        End If
        If sel_pathogen_unit_rs("MID_NAME") <> "" Then
            mname = sel_pathogen_unit_rs("MID_NAME")
        End If
        If sel_pathogen_unit_rs("LAST_NAME") <> "" Then
            lname = sel_pathogen_unit_rs("LAST_NAME")
        End If
        pname = Trim(sal) & " " & Trim(FNAME) & " " & Trim(mname) & " " & Trim(lname)
             
        If Len(Trim(sel_pathogen_unit_rs("BIRTH_DATE"))) <> 0 Then
            age = DateDiff("YYYY", sel_pathogen_unit_rs("BIRTH_DATE"), CStr(DATE11))
        End If
          
        If MSFlexGrid1.TextMatrix(i, 1) <> "" Then
            MSFlexGrid1.TextMatrix(i, 6) = Trim(sel_pathogen_unit_rs("SEX"))
        End If
        
        If MSFlexGrid1.TextMatrix(i, 1) <> "" Then
            MSFlexGrid1.TextMatrix(i, 7) = age
        End If
        
         '***** Shraddha 10/10/2011 *****
        If Trim(sel_pathogen_unit_rs("ORGSO1")) <> "" Or Trim(sel_pathogen_unit_rs("ORGSO2")) <> "" Or Trim(sel_pathogen_unit_rs("ORGSO3")) <> "" Then
            patho_grown = "  Pathogen grown : Yes"
        Else
            patho_grown = "  Pathogen grown : No"
        End If
        '*******************************
                
        If i <> 1 Then
            If Trim(UCase(SP)) = Trim(UCase(sel_pathogen_unit_rs("SPECIMEN"))) And cno = sel_pathogen_unit_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 8) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 8) = Trim(sel_pathogen_unit_rs("SPECIMEN")) & patho_grown
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 8) = Trim(sel_pathogen_unit_rs("SPECIMEN")) & patho_grown
        End If
        SP = Trim(sel_pathogen_unit_rs("SPECIMEN"))
        
        If i <> 1 Then
            If Trim(colldt) = Format(sel_pathogen_unit_rs("samcoldate"), "dd/mm/yyyy") And cno = sel_pathogen_unit_rs("caseno") Then
                MSFlexGrid1.TextMatrix(i, 9) = ""
            Else
                MSFlexGrid1.TextMatrix(i, 9) = Format(sel_pathogen_unit_rs("samcoldate"), "dd/mm/yyyy")
            End If
        Else
                MSFlexGrid1.TextMatrix(i, 9) = Format(sel_pathogen_unit_rs("samcoldate"), "dd/mm/yyyy")
        End If
        If (sel_pathogen_unit_rs("samcoldate")) <> Null Or (sel_pathogen_unit_rs("samcoldate")) <> "" Then
            colldt = Format(sel_pathogen_unit_rs("samcoldate"), "dd/mm/yyyy")
        End If
                
        If Trim(sel_pathogen_unit_rs("ORGSO1")) <> "" Then
            If i <> 1 Then
                If Trim(or1) = Trim(sel_pathogen_unit_rs("ORGSO1")) And cno = sel_pathogen_unit_rs("caseno") Then
                    MSFlexGrid1.TextMatrix(i, 10) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 10) = sel_pathogen_unit_rs("ORGSO1")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 10) = sel_pathogen_unit_rs("ORGSO1")
            End If
            or1 = sel_pathogen_unit_rs("ORGSO1")
        End If

        If Trim(sel_pathogen_unit_rs("ORGSO1")) <> "" Then
            If Trim(sel_pathogen_unit_rs("ANTIBIO1")) <> "" Then
                If i <> 1 Then
                    If Trim(antbt1) = Trim(sel_pathogen_unit_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO1")) & ")" And cno = sel_pathogen_unit_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 11) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 11) = sel_pathogen_unit_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO1")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 11) = sel_pathogen_unit_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO1")) & ")"
                End If
                antbt1 = Trim(sel_pathogen_unit_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO1")) & ")"
                
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 11
                If Trim(sel_pathogen_unit_rs("ANTIBIO1")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_unit_rs("ANTIBIO1")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_unit_rs("ANTIBIO1")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
            End If
        End If
        
        If Trim(sel_pathogen_unit_rs("ORGSO2")) <> "" Then
            If i <> 1 Then
                If Trim(or2) = Trim(sel_pathogen_unit_rs("ORGSO2")) And cno = sel_pathogen_unit_rs("caseno") Then
                    MSFlexGrid1.TextMatrix(i, 12) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 12) = sel_pathogen_unit_rs("ORGSO2")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 12) = sel_pathogen_unit_rs("ORGSO2")
            End If
            or2 = sel_pathogen_unit_rs("ORGSO2")
        End If
        
        If Trim(sel_pathogen_unit_rs("ORGSO2")) <> "" Then
            If Trim(sel_pathogen_unit_rs("ANTIBIO2")) <> "" Then
                If i <> 1 Then
                    If Trim(antbt2) = Trim(sel_pathogen_unit_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO2")) & ")" And cno = sel_pathogen_unit_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 13) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 13) = sel_pathogen_unit_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO2")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 13) = sel_pathogen_unit_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO2")) & ")"
                End If
                antbt2 = Trim(sel_pathogen_unit_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO2")) & ")"
                
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 13
                If Trim(sel_pathogen_unit_rs("ANTIBIO2")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_unit_rs("ANTIBIO2")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_unit_rs("ANTIBIO2")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
            End If
        End If
        
        If Trim(sel_pathogen_unit_rs("ORGSO3")) <> "" Then
            If i <> 1 Then
                If Trim(or3) = Trim(sel_pathogen_unit_rs("ORGSO3")) And cno = sel_pathogen_unit_rs("caseno") Then
                    MSFlexGrid1.TextMatrix(i, 14) = ""
                Else
                    MSFlexGrid1.TextMatrix(i, 14) = sel_pathogen_unit_rs("ORGSO3")
                End If
            Else
                MSFlexGrid1.TextMatrix(i, 14) = sel_pathogen_unit_rs("ORGSO3")
            End If
            or3 = sel_pathogen_unit_rs("ORGSO3")
        End If

        If Trim(sel_pathogen_unit_rs("ORGSO3")) <> "" Then
            If Trim(sel_pathogen_unit_rs("ANTIBIO3")) <> "" Then
                If i <> 1 Then
                    If Trim(antbt3) = Trim(sel_pathogen_unit_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO3")) & ")" And cno = sel_pathogen_unit_rs("caseno") Then
                        MSFlexGrid1.TextMatrix(i, 15) = ""
                    Else
                        MSFlexGrid1.TextMatrix(i, 15) = sel_pathogen_unit_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO3")) & ")"
                    End If
                Else
                    MSFlexGrid1.TextMatrix(i, 15) = sel_pathogen_unit_rs("ANTIBIOTIC") & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO3")) & ")"
                End If
                antbt3 = Trim(sel_pathogen_unit_rs("ANTIBIOTIC")) & "(" & Trim(sel_pathogen_unit_rs("ANTIBIO3")) & ")"
                
                MSFlexGrid1.row = i
                MSFlexGrid1.col = 15
                If Trim(sel_pathogen_unit_rs("ANTIBIO3")) = "S" Then
                    MSFlexGrid1.CellBackColor = &H80FF80
                End If
                If Trim(sel_pathogen_unit_rs("ANTIBIO3")) = "I" Then
                    MSFlexGrid1.CellBackColor = &H80FFFF
                End If
                If Trim(sel_pathogen_unit_rs("ANTIBIO3")) = "R" Then
                    MSFlexGrid1.CellBackColor = &H8080FF
                End If
            End If
        End If

        cno = sel_pathogen_unit_rs("caseno")
        sel_pathogen_unit_rs.MoveNext
    Next


End If
'*******************************
End Sub


Public Sub proc1()
If title_chk = True Then
    lbl_caption = "No records found"
    MousePointer = vbNormal
    Exit Sub
End If

If Option1.Value = True Then
    If Opt_specimen.Value = True Then
        wardname = Trim(cmb_ward.Text)
        specimen = Trim(cmb_specimen.Text)
        Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " which belong to " & wardname & " with Specimen : " & specimen
    End If
    
    If Opt_antibiotic.Value = True Then
        wardname = Trim(cmb_ward.Text)
        antibiotic = Trim(cmb_antibiotic.Text)
        Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " which belong to " & wardname & " with Antiboitic : " & antibiotic
    End If
    
    '***** Shraddha 09/12/09 *****
    If Opt_organism.Value = True Then
        wardname = Trim(cmb_ward.Text)
        organism = Trim(cmb_organism.Text)
        Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " which belong to " & wardname & " with Antiboitic : " & organism
    End If
    '*****************************
    
    If Opt_specimen.Value = False And Opt_antibiotic.Value = False And Opt_organism.Value = False Then
        Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " which belong to " & wardname
    End If
End If

If Option2.Value = True Then
    If Opt_specimen.Value = True Then
        specimen = Trim(cmb_specimen.Text)
        Title = "All ICU records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " with Specimen : " & specimen
    End If
    
    If Opt_antibiotic.Value = True Then
        antibiotic = Trim(cmb_antibiotic.Text)
        Title = "All ICU records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " with Antiboitic : " & antibiotic
    End If
    
    '***** Shraddha 09/12/09 *****
    If Opt_organism.Value = True Then
        organism = Trim(cmb_organism.Text)
        Title = "All ICU records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " with Organism : " & organism
    End If
    '*****************************
    
    If Opt_specimen.Value = False And Opt_antibiotic.Value = False And Opt_organism.Value = False Then
        Title = "All ICU records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText)
    End If
    
End If

If Option3.Value = True Then
    c_no = MaskEdBox3.FormattedText
    Title = "All records of " & c_no & " between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText)
End If

If Option4.Value = True Then
    If cmb_unit.Visible = True Then
        unit = Trim(cmb_unit.Text)
        Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " with Unit/Doctor : " & unit
    End If
    If cmb_doctor.Visible = True Then
        doctor = Trim(cmb_doctor.Text)
        unit = Trim(cmb_unit.Text)
        Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " with Unit : " & unit & " and doctor : " & doctor
    End If
End If

If opt_OP.Value = True Then
    Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " who are Outpatient"
End If

lbl_caption.Caption = Title

MousePointer = vbNormal
End Sub

Private Sub cmb_unit_LostFocus()
Dim chk_unit_qr As rdoQuery
Dim chk_unit_rs As rdoResultset
Dim sql_chk_unit As String

sql_chk_unit = "SELECT U_D FROM PATOOLSLIB.MA1031 where UNIT_NAME=? and U_D='U'"
Set chk_unit_qr = rdb.CreateQuery("", sql_chk_unit)

chk_unit_qr(0) = Trim(cmb_unit.Text)
Set chk_unit_rs = chk_unit_qr.OpenResultset(1, 3)

If Not chk_unit_rs.EOF Then
    
    cmb_doctor.clear
    '*************************** doctor ***************************
    'SELECT * FROM PATOOLSLIB/MA1031 163
    'UNIT_NAME  57 D=105
    Dim doctor_QR As rdoQuery
    Dim doctor_RS As rdoResultset
    Dim sql_doctor As String
    
    sql_doctor = "SELECT UD_NAME FROM PATOOLSLIB.MA1031 where UNIT_NAME=? and U_D='D'"
    Set doctor_QR = rdb.CreateQuery("", sql_doctor)
    
    doctor_QR(0) = Trim(cmb_unit.Text)
    Set doctor_RS = doctor_QR.OpenResultset(1, 3)
    
    If Not doctor_RS.EOF Then
        lbl_doctor.Visible = True
        cmb_doctor.Visible = True
        
        While Not doctor_RS.EOF
            If Trim(doctor_RS(0)) <> "" Then
                cmb_doctor.AddItem doctor_RS(0)
            End If
            doctor_RS.MoveNext
        Wend
    Else
    
        lbl_doctor.Visible = False
        cmb_doctor.Visible = False
            
    End If
    doctor_RS.Close
    '***********************************************************************
Else
    lbl_doctor.Visible = False
    cmb_doctor.Visible = False
End If
End Sub



Public Sub InsDataPathogen()
    Dim rq_pathogen1 As rdoQuery
    Dim rs_distpathogen As rdoResultset
    Dim rs_pathogen As rdoResultset
    Dim rs_dmicbantib As rdoResultset
    
    '''FOR PATHOGEN2
    
    Dim rq_pathogen2 As rdoQuery
    Dim rs_pathogen2 As rdoResultset
    Dim rs_dmicbantib2 As rdoResultset
    
    rdb.BeginTrans
   ' rdb.Execute ("delete  from dislib.pathogen1")
    'rdb.Execute ("delete from dislib.pathogen2")
    'rdb.Execute ("delete  from dislib.pathogen3")
    
    MSFlexGrid1.clear
    MSFlexGrid1.Cols = 225
    MSFlexGrid1.Visible = True
    
    Dim j As Integer
    Dim row As Long
    row = 0
   ' Set rq_pathogen1 = rdb.CreateQuery("", "insert into dislib.pathogen1 values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")
    If Option1.Value = True And cmb_ward.Text = "ALL" Then
        If Opt_specimen.Value = True Then
            If cmb_specimen.Text = "ALL" Then
                Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where reqndt between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            Else
                Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where specimen='" & Trim(cmb_specimen.Text) & "' reqndt between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            End If
            Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " with Specimen : " & Trim(cmb_specimen.Text)
        End If
        If Opt_organism.Value = True Then
            If cmb_organism.Text = "All" Then
                Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where reqndt between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            Else
                Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where (ORGSO1='" & Trim(cmb_organism.Text) & "' or orgso2='" & Trim(cmb_organism.Text) & "' or orgso3='" & Trim(cmb_organism.Text) & "') and reqndt between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            End If
            Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " with Organism : " & Trim(cmb_organism.Text)
        End If
        If Opt_antibiotic.Value = True Then
            Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where reqndt between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " with Antibiotic : " & Trim(cmb_antibiotic.Text)
        End If
    ElseIf Option1.Value = True And cmb_ward.Text <> "ALL" Then
        If Opt_specimen.Value = True Then
            If cmb_specimen.Text = "ALL" Then
                Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where WARD_NAME='" & Trim(cmb_ward.Text) & "' and  reqndt between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            Else
                Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where WARD_NAME='" & Trim(cmb_ward.Text) & "' and specimen='" & Trim(cmb_specimen.Text) & "' AND reqndt between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            End If
            Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " which belong to " & cmb_ward.Text & " with Specimen: " & Trim(cmb_specimen.Text)
        End If
        If Opt_organism.Value = True Then
            If cmb_organism.Text = "ALL" Then
                Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where ward_name='" & Trim(cmb_ward.Text) & "' and  reqndt between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            Else
                Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where ward_name='" & Trim(cmb_ward.Text) & "' and  (ORGSO1='" & Trim(cmb_organism.Text) & "' or orgso2='" & Trim(cmb_organism.Text) & "' or orgso3='" & Trim(cmb_organism.Text) & "') and reqndt between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            End If
            Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " which belong to " & cmb_ward.Text & " with Organism: " & Trim(cmb_organism.Text)
        End If
        If Opt_antibiotic.Value = True Then
            Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where ward_name='" & Trim(cmb_ward.Text) & "' and  reqndt  between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " which belong to " & cmb_ward.Text & " with Antibiotic: " & Trim(cmb_antibiotic.Text)
        End If
    End If
    If Option2.Value = True Then
        If Opt_specimen.Value = True Then
            If cmb_specimen.Text = "ALL" Then
                Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where ICDS2='YES' and  reqndt between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            Else
                Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where ICDS2='YES' and specimen='" & Trim(cmb_specimen.Text) & "' AND reqndt between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            End If
            Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " which belong to ICU with Specimen: " & Trim(cmb_specimen.Text)
        End If
        If Opt_organism.Value = True Then
            If cmb_organism.Text = "All" Then
                Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where ICDS2='YES' and  reqndt between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            Else
                Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where ICDS2='YES' and  (ORGSO1='" & Trim(cmb_organism.Text) & "' or orgso2='" & Trim(cmb_organism.Text) & "' or orgso3='" & Trim(cmb_organism.Text) & "') and reqndt between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
            End If
            Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " which belong to ICU with Organism: " & Trim(cmb_organism.Text)
        End If
        If Opt_antibiotic.Value = True Then
            Set rs_distpathogen = rdb.OpenResultset("select distinct(reqnno) from rislib.pathogen where ICDS2='YES' and  reqndt  between '" & Format(MaskEdBox1.FormattedText, "yyyy-mm-dd") & "' and '" & Format(MaskEdBox2.FormattedText, "yyyy-mm-dd") & "'", 1, 1)
        End If
            Title = "All records between " & Trim(MaskEdBox1.FormattedText) & " and " & Trim(MaskEdBox2.FormattedText) & " which belong to ICU with Antibiotic: " & Trim(cmb_antibiotic.Text)
    End If
    lbl_caption.Caption = Title
    
    If Not rs_distpathogen.EOF Then
        Do While Not rs_distpathogen.EOF
        ''''***** FOR INSERTING IN PATOHOGEN1******
            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
            row = row + 1
            With MSFlexGrid1
            Set rs_pathogen = rdb.OpenResultset("select caseno,SAMCOLDATE,ANTIBIO1,antibiotic,orgso1,SPECIMEN from rislib.pathogen where reqnNo='" & Trim(rs_distpathogen(0)) & "'", 1, 1)
            If Not rs_pathogen.EOF Then
                
                'rq_pathogen1(0) = rs_distpathogen(0)
                .TextMatrix(row, 0) = rs_distpathogen(0)
                'rq_pathogen1(1) = Trim(rs_pathogen(0))
                .TextMatrix(row, 1) = Trim(rs_pathogen(0))
                If IsNull(rs_pathogen(1)) = True Then
                'rq_pathogen1(2) = Null
                .TextMatrix(row, 2) = ""
                Else
                'rq_pathogen1(2) = rs_pathogen(1)
                .TextMatrix(row, 2) = rs_pathogen(1)
                End If
                'rq_pathogen1(3) = rs_pathogen(4)
                .TextMatrix(row, 3) = rs_pathogen(5) '''SPECIMEN
                           
                .TextMatrix(row, 4) = rs_pathogen(4) '''ORGSO1
                For i = 5 To 76
                   ' rq_pathogen1(i) = ""
                    .TextMatrix(row, i) = ""
                Next
                Do While Not rs_pathogen.EOF
                    Set rs_dmicbantib = rdb.OpenResultset("select ANTIBCODE from dislib.dmicbantib where MODALITY='BACTERIOLOGY' AND ANTIBDESC='" & Trim(rs_pathogen("antibiotic")) & "'", 1, 1)
                    j = Trim(rs_dmicbantib(0)) + 4
                    'rq_pathogen1(j) = rs_pathogen(2)
                    .TextMatrix(row, j) = rs_pathogen(2)
                
                    rs_pathogen.MoveNext
                Loop
               ' rq_pathogen1.Execute
            End If
            
        '''***** FOR INSERTING IN PATOHOGEN2******
       ' Set rq_pathogen2 = rdb.CreateQuery("", "insert into dislib.pathogen2 values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")
            Set rs_pathogen2 = rdb.OpenResultset("select caseno,SAMCOLDATE,ANTIBIO2,antibiotic,orgso2 from rislib.pathogen where reqnNo='" & Trim(rs_distpathogen(0)) & "'", 1, 1)
            If Not rs_pathogen2.EOF Then
               ' rq_pathogen2(0) = rs_distpathogen(0)
               ' rq_pathogen2(1) = Trim(rs_pathogen2(0))
                If IsNull(rs_pathogen2(1)) = True Then
                  '  rq_pathogen2(2) = Null
                Else
                   ' rq_pathogen2(2) = rs_pathogen2(1)
                End If
                If IsNull(rs_pathogen2(4)) = False Then
                   ' rq_pathogen2(3) = rs_pathogen2(4)
                    .TextMatrix(row, 77) = rs_pathogen2(4)
                Else
                   ' rq_pathogen2(3) = Null
                    .TextMatrix(row, 77) = ""
                End If
                For i = 5 To 76
                   ' rq_pathogen2(i) = ""
                    .TextMatrix(row, i + 74) = ""
                Next
                Do While Not rs_pathogen2.EOF
                    Set rs_dmicbantib2 = rdb.OpenResultset("select ANTIBCODE from dislib.dmicbantib where MODALITY='BACTERIOLOGY' AND  ANTIBDESC='" & Trim(rs_pathogen2("antibiotic")) & "'", 1, 1)
                    j = Trim(rs_dmicbantib2(0)) + 3
                    If IsNull(rs_pathogen2(2)) = False Then
                   ' rq_pathogen2(j) = rs_pathogen2(2)
                    .TextMatrix(row, j + 74) = rs_pathogen2(2)
                    Else
                    'rq_pathogen2(j) = ""
                    .TextMatrix(row, j + 74) = ""
                    End If
                    rs_pathogen2.MoveNext
                Loop
                'rq_pathogen2.Execute
            End If
           ''''UPTO HERE'''''''''
             '''***** FOR INSERTING IN PATOHOGEN3******
       ' Set rq_pathogen2 = rdb.CreateQuery("", "insert into dislib.pathogen2 values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")
            Set rs_pathogen3 = rdb.OpenResultset("select caseno,SAMCOLDATE,ANTIBIO3,antibiotic,orgso3 from rislib.pathogen where reqnNo='" & Trim(rs_distpathogen(0)) & "'", 1, 1)
            If Not rs_pathogen3.EOF Then
               ' rq_pathogen2(0) = rs_distpathogen(0)
               ' rq_pathogen2(1) = Trim(rs_pathogen2(0))
                If IsNull(rs_pathogen3(1)) = True Then
                  '  rq_pathogen2(2) = Null
                Else
                   ' rq_pathogen2(2) = rs_pathogen2(1)
                End If
                If IsNull(rs_pathogen3(4)) = False Then
                   ' rq_pathogen2(3) = rs_pathogen2(4)
                    .TextMatrix(row, 151) = rs_pathogen3(4)
                Else
                   ' rq_pathogen2(3) = Null
                    .TextMatrix(row, 151) = ""
                End If
                For i = 4 To 76
                   ' rq_pathogen2(i) = ""
                    .TextMatrix(row, i + 148) = ""
                Next
                Do While Not rs_pathogen3.EOF
                    Set rs_dmicbantib3 = rdb.OpenResultset("select ANTIBCODE from dislib.dmicbantib where MODALITY='BACTERIOLOGY' AND  ANTIBDESC='" & Trim(rs_pathogen3("antibiotic")) & "'", 1, 1)
                    j = Trim(rs_dmicbantib3(0)) + 3
                    If IsNull(rs_pathogen3(2)) = False Then
                   ' rq_pathogen2(j) = rs_pathogen2(2)
                    .TextMatrix(row, j + 148) = rs_pathogen3(2)
                    Else
                    'rq_pathogen2(j) = ""
                    .TextMatrix(row, j + 148) = ""
                    End If
                    rs_pathogen3.MoveNext
                Loop
                'rq_pathogen2.Execute
            End If
           ''''UPTO HERE'''''''''
            End With
            rs_distpathogen.MoveNext
            
        Loop
    End If
    
    
    
    
    
End Sub
