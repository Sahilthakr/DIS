VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form dept_collection_BARCODE2 
   Caption         =   "Department Sample Receiving  Centre"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   120
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   120
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
      Left            =   10200
      TabIndex        =   19
      Top             =   840
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
      Left            =   10200
      TabIndex        =   18
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame buttons 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
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
      TabIndex        =   16
      Top             =   7560
      Width           =   11655
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
         Left            =   10320
         TabIndex        =   26
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
         Left            =   9000
         MouseIcon       =   "dept_collection_BARCODE.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   10
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
         Left            =   7560
         MouseIcon       =   "dept_collection_BARCODE.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   9
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
         Left            =   6120
         MouseIcon       =   "dept_collection_BARCODE.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   7
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
         Left            =   4800
         MouseIcon       =   "dept_collection_BARCODE.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   6
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
         Left            =   3360
         MouseIcon       =   "dept_collection_BARCODE.frx":0C28
         MousePointer    =   99  'Custom
         TabIndex        =   17
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
         MouseIcon       =   "dept_collection_BARCODE.frx":0F32
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_new 
         Caption         =   "&New"
         Default         =   -1  'True
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
         MouseIcon       =   "dept_collection_BARCODE.frx":123C
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
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
      Height          =   4455
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   11655
      Begin VB.CommandButton cmd_VPP 
         Caption         =   "VPP"
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
         Left            =   10200
         TabIndex        =   28
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CMB_PRINT1 
         Caption         =   "&Print"
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
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txt_pathno 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   10200
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chk_collected 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Collected"
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
         Left            =   6600
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chk_rcvd 
         BackColor       =   &H8000000E&
         Caption         =   "Received"
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
         Left            =   8280
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5953
         _Version        =   393216
         Rows            =   3
         Cols            =   7
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         ForeColorFixed  =   192
         AllowUserResizing=   3
         FormatString    =   $"dept_collection_BARCODE.frx":1546
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
      Begin VB.Label LBL_VPP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Letter for Vpp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   8760
         TabIndex        =   29
         Top             =   4080
         Visible         =   0   'False
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
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
      Height          =   1335
      Left            =   480
      TabIndex        =   12
      Top             =   1320
      Width           =   10815
      Begin VB.TextBox Text1 
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
         Left            =   1800
         TabIndex        =   32
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox TXT_NAME 
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
         Left            =   5520
         TabIndex        =   24
         Top             =   240
         Width           =   4455
      End
      Begin MSMask.MaskEdBox msk_caseno 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
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
      Begin VB.ComboBox Cmb_modality 
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
         Left            =   8280
         TabIndex        =   2
         Top             =   840
         Width           =   2415
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
         Left            =   4680
         TabIndex        =   1
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Case No :-"
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
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisation No :-"
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
         Width           =   1500
      End
      Begin VB.Label LBL_NAME 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :-"
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
         Left            =   4680
         TabIndex        =   23
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lbl_mod 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modality :-"
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
         Left            =   7320
         TabIndex        =   15
         Top             =   840
         Width           =   900
      End
      Begin VB.Label lbl_dept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department :-"
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
         Left            =   3360
         TabIndex        =   14
         Top             =   840
         Width           =   1170
      End
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
      Left            =   960
      TabIndex        =   30
      Top             =   120
      Width           =   8595
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
      Left            =   9600
      TabIndex        =   21
      Top             =   840
      Width           =   420
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
      Left            =   9600
      TabIndex        =   20
      Top             =   480
      Width           =   420
   End
   Begin VB.Label lbl_req 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT SAMPLE RECEIVING WITH BARCODE "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   1080
      TabIndex        =   11
      Top             =   720
      Width           =   7995
   End
End
Attribute VB_Name = "dept_collection_BARCODE2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SELMA106_QR As rdoQuery
Dim SELMA106_RS As rdoResultset
Dim SELMA106FLAG As Boolean
Dim SELRMODALITY_QR As rdoQuery
Dim SELRMODALITY_RS As rdoResultset
Dim SELRMODALITYFLAG As Boolean
Dim deptcollectionflag As Boolean
Dim NEWFLAG As Boolean
Dim viewflag As Boolean
Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG As Boolean
Dim REC_DATE1, REC_TIME1
Dim SELRREQUISITN_QR1 As rdoQuery
Dim SELRREQUISITN_RS1 As rdoResultset
Dim SELRREQUISITNFLAG1 As Boolean
'INTERFACE
'ADDED BY AMIT FOR RTIMELOG
'Dim INSTIME_QR As rdoQuery
'Dim INSTIME_RS As rdoResultset


Dim SELPATH_QR As rdoQuery
Dim SELPATH_RS As rdoResultset
Dim SELPROCODE_QR As rdoQuery
Dim SELPROCODE_RS As rdoResultset
Dim LABNO As Variant

Public R1, C1, II
Dim BARCODE_NEW As Boolean
Dim BARCODE_VIEW As Boolean
Dim SELLABNO_QR As rdoQuery
Dim SELLABNO_RS As rdoResultset
Public REQNO As Variant
Public CASENO As Variant
Dim dept As String
Dim MODALITY As String
Dim SELMA101_QR As rdoQuery
Dim SELMA101_RS As rdoResultset
Dim SELMA101FLAG As Boolean
Dim updrrequisitn_qr As rdoQuery
Dim MODIFYFLAG As Boolean
Dim SELRREQNO_QR As rdoQuery
Dim SELRREQNO_RS As rdoResultset
Dim SELRREQNOFLAG As Boolean

Dim SELRREQNO_QR1 As rdoQuery
Dim SELRREQNO_RS1 As rdoResultset
Dim SELRREQNOFLAG1 As Boolean

Dim SELPATHNO_QR As rdoQuery
Dim SELPATHNO_RS As rdoResultset

Dim SELTA1011_QR As rdoQuery
Dim SELTA1011_RS As rdoResultset
Dim SELTA1011FLAG As Boolean
Dim SELDREQSTAFFD1_QR As rdoQuery
Dim SELDREQSTAFFD1_RS As rdoResultset
Dim SELDREQSTAFFD1FLAG As Boolean
Dim upddrequisitn As rdoQuery
Dim SELDREQUISITNQR As rdoQuery
Dim SELDREQUISITNRS As rdoResultset
Dim seldequisitnflag As Boolean
Dim SELDREQQR As rdoQuery
Dim SELDREQRS As rdoResultset
Dim SELDREQFLAG As Boolean
Dim SELREQNOQR As rdoQuery
Dim SELREQNORS As rdoResultset
Dim SELREQNORSFLAG As Boolean
Dim DEPTMODALITY As String
Dim INSDBCINFACEQR As rdoQuery

Dim SELDMULPROCQR As rdoQuery
Dim seldmulprocrs As rdoResultset
Dim seldmulprocflag As Boolean

Dim SELELPROSUBQR As rdoQuery
Dim SELELPROSUBRS As rdoResultset
Dim SELELPROSUBFLAG As Boolean
Dim srow As Integer

''DTMNO

Dim SELDTMNO_QR As rdoQuery
Dim SELDTMNO_RS As rdoResultset
Dim SELDTMNO_FLAG As Boolean

''STAFF
Dim SELEMPBASINFO_QR As rdoQuery
Dim SELEMPBASINFO_RS As rdoResultset
Dim SELEMPBASINFO_FLAG As Boolean


'''RECEIPT TIME
Dim UPDLATERALITS_QR As rdoQuery

Dim SELDAUTOPATH_QR As rdoQuery
Dim SELDAUTOPATH_RS As rdoResultset
Dim SELDAUTOPATHFLAG As Boolean

Dim SELTMAUTOPATH_QR As rdoQuery
Dim SELTMAUTOPATH_RS As rdoResultset
Dim SELTMAUTOPATHFLAG As Boolean

Dim DPATHYEAR_QR As rdoQuery
Dim DPATHYEAR_RS As rdoResultset
Dim DPATHYEARFLAG As Boolean

Dim UPDDAUTOPATH_QR As rdoQuery
Dim INSDAUTOPATH_QR As rdoQuery

Dim UPDTMAUTOPATH_QR As rdoQuery
Dim INSTMAUTOPATH_QR As rdoQuery

Dim SELRREQUISITN_QR11 As rdoQuery
Dim SELRREQUISITN_RS11 As rdoResultset
Dim SELRREQUISITN_FLAG11 As Boolean

Dim SELRREQNO_QR11 As rdoQuery
Dim SELRREQNO_RS11 As rdoResultset
Dim SELRREQNO_FLAG11 As Boolean
'SAVE RECEIVING DATE AND TIME IN RTIMELOG BY AMIT
Dim UPDTIME_QR As rdoQuery
Dim UPDTIME_RS As rdoResultset
Dim SELTIME_QR As rdoQuery
Dim SELTIME_RS As rdoResultset
Dim SELTIME_FLG As Boolean

Dim pathno As String

''*********tracking System*************
Dim Ins_Spathology_qr As rdoQuery
''*********tracking System*************
'Public CASENO
Public PATNAME
Public COLL_CENTER
Public REC_DATE
Public REC_TIME

Dim SELDREQUISITN_PATH_QR As rdoQuery
Dim SELDREQUISITN_PATH_RS As rdoResultset
Dim SELDREQUISITN_PATH_flg As Boolean

''****INTERFACE CODE BY AMIT *****************
Dim INSDINTERFACE_QR As rdoQuery
Dim INSBINTERFACE_QR As rdoQuery
Dim RS_DMULPROCE  As rdoResultset


Dim seldmulproce_qr As rdoQuery
Dim seldmulproce_rs As rdoResultset
Dim SELDMULPROCE_FLAG As Boolean
Dim scrlck_flg As Boolean
Dim SELSAMPLECODE_QR As rdoQuery
Dim SELSAMPLECODE_RS As rdoResultset

'' HEMATO INTERFACE
Dim INSHINTERFACE_QR As rdoQuery
Dim INSCINTERFACE_QR As rdoQuery
Dim SELDHEMCYTO_QR As rdoQuery
Dim SELDHEMCYTO_RS As rdoResultset
'Dim seldmulproce_qr As rdoQuery
'Dim seldmulproce_rs As rdoResultset
Dim seldhlprosub_qr As rdoQuery
Dim seldhlprosub_rs As rdoResultset

'' HEMATO INTERFACE

''****INTERFACE CODE BY AMIT AGRAWAL*****************
' tamankar begins
Dim interexpcd_qr As rdoQuery
Dim interexpxd_rs As rdoResultset
Dim INTERFACEDATA_QR0 As rdoQuery
Dim INTERFACEDATA_RS0 As rdoResultset
Dim INTERFACEDATA_QR1 As rdoQuery
Dim INTERFACEDATA_RS1 As rdoResultset
Dim SELDBCPROSUB_QR As rdoQuery
Dim SELDBCPROSUB_RS As rdoResultset

Private Sub chk_collected_Click()
'MSFlexGrid1.Col = 4
'If chk_collected.Value = 1 Then
'    MSFlexGrid1.Text = "COLLECTED"
'Else
'    MSFlexGrid1.Text = ""
'End If

End Sub

Private Sub chk_collected_KeyPress(KeyAscii As Integer)
On Error GoTo MSG:
'If KeyAscii = 13 Then
'If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
'        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
'        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
'End If
'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
'
'        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'        chk_collected = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4))
'        chk_collected.SetFocus
'        chk_collected.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'
'End If
'End If
'Exit Sub
MSG:
'MsgBox Err.Description
End Sub

Private Sub chk_rcvd_Click()
Dim num As String
'If dept = "NZZ" Then
'    MSFlexGrid1.Col = 5
'    If chk_rcvd.Value = 1 Then
'        MSFlexGrid1.Text = "RECEIVED"
'    Else
'        MSFlexGrid1.Text = ""
'    End If
'    Exit Sub
'End If

If MODALITY = "TM" And chk_rcvd.Value = 1 And Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))) = 0 Then
    'justin 23-08-2004
    'i = Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1), 11, 6)
    'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = "TM/" & i
    '    num = Format(Now, "yymmdd")
    '    SELDREQUISITN_PATH_QR(0) = num & "%"
    '
    '    If SELDREQUISITN_PATH_flg = False Then
    '       Set SELDREQUISITN_PATH_RS = SELDREQUISITN_PATH_QR.OpenResultset(2, 3)
    '       SELDREQUISITN_PATH_flg = True
    '
    '       If IsNull(SELDREQUISITN_PATH_RS(0)) = False Then
    '          MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = SELDREQUISITN_PATH_RS(0) + 1
    '          pathno = SELDREQUISITN_PATH_RS(0) + 1
    '       Else
    '          MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = num & "001"
    '          pathno = num & "001"
    '       End If
    '    Else
    '     pathno = "0" & Val(pathno) + 1
    '     MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = pathno
    '    End If
    'JUSTIN

ElseIf Trim(MODALITY) = "TM" And chk_rcvd.Value = 0 Then
    '   If Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))) <> 0 And (Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)) = pathno) Then
    '    pathno = "0" & Val(pathno) - 1
    '   End If
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = " "
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = " "
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = ""
End If
MSFlexGrid1.col = 5
If chk_rcvd.Value = 1 Then
    MSFlexGrid1.Text = "RECEIVED"
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = "COLLECTED"
    'chk_collected.Value = 1
Else
    MSFlexGrid1.Text = ""
End If

If Trim(dept) = "FZZ" And Trim(MODALITY) = "SP" And chk_rcvd.Value = 1 Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = "RECEIVED"
    'SPAUTONO
ElseIf Trim(dept) = "FZZ" And Trim(MODALITY) = "SP" And chk_rcvd.Value = 0 Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = ""
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
End If

If Trim(dept) = "UZZ" And chk_rcvd.Value = 1 Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = "RECEIVED"
    'SPAUTONO
ElseIf Trim(dept) = "UZZ" Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = ""
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
End If

End Sub

Private Sub chk_rcvd_KeyPress(KeyAscii As Integer)
On Error GoTo MSG:
If KeyAscii = 13 Then
    If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
    End If

    If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.row = MSFlexGrid1.row + 1
        chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
        chk_rcvd.SetFocus
        chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    End If
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

'Private Sub chk_rcvd_LostFocus()
'txt_pathno.SetFocus
'End Sub

Private Sub cmb_dept_Click()
Dim i As Integer
On Error GoTo MSG:
i = InStr(Trim(cmb_dept.Text), "$")
dept = Mid(cmb_dept.Text, i + 1, 4)
If SELRMODALITYFLAG = False Then
    SELRMODALITY_QR(0) = Trim(dept)
    'DEPT = Trim(Mid(Trim(cmb_dept.Text), i + 1, 4))
    Set SELRMODALITY_RS = SELRMODALITY_QR.OpenResultset(1, 3)
Else
    SELRMODALITY_QR(0) = Trim(dept)
    SELRMODALITY_RS.MoreResults
    SELRMODALITY_RS.Requery
End If

If SELRMODALITY_RS.RowCount > 0 Then
    cmb_modality.clear
    While Not SELRMODALITY_RS.EOF
        cmb_modality.AddItem SELRMODALITY_RS("modlitydes") & "$" & SELRMODALITY_RS("modalitycd")
        SELRMODALITY_RS.MoveNext
    Wend
End If
If dept = "NZZ" Then
    cmb_modality.AddItem "ALL"
End If
cmb_modality.Enabled = True
cmb_modality.SetFocus

If dept = "FEC" And viewflag = True Then
  cmd_modify.Enabled = True
End If

Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmb_dept_GotFocus()
On Error GoTo MSG:
Dim ind_cbodept As Integer
If Trim(cmb_dept) = "" Then
    For ind_cbodept = 0 To (cmb_dept.ListCount - 1)
        If Trim(cmb_dept.List(ind_cbodept)) = "PATHOLOGY$FZZ" Then
            cmb_dept.ListIndex = ind_cbodept
        End If
    Next ind_cbodept
End If
Exit Sub
MSG:
MsgBox Error(Err)
End Sub

Private Sub cmb_modality_Click()
Dim i As Integer
i = InStr(cmb_modality, "$")
MODALITY = Mid(cmb_modality, i + 1, 4)
''If modality = "IH" Then
''dept = "FDZ"
''End If
''msk_fromdt.Text = Format(DATE11, "dd/mm/yyyy")
''msk_todt.Text = Format(DATE11, "dd/mm/yyyy")
''msk_fromdt.SetFocus
MSFlexGrid1.FormatString = " Sr.No |^   Requisation No.   |^   Case No.   |^    Procedure Name          |^  Sample Collected  |^    Sample Received |^    PATH No. "
If MODALITY = "BC" Or MODALITY = "EL" Or MODALITY = "TM" Or MODALITY = "BA" Or MODALITY = "XX" Or MODALITY = "SE" Or MODALITY = "MB" Or MODALITY = "MC" Or MODALITY = "CM" Or MODALITY = "MO" Then
    MSFlexGrid1.FormatString = " Sr.No |^   Requisition No.   |^   Case No.   |^    Procedure Name          |^  Sample Collected  |^    Sample Received |^    LAB No.   "
End If
If MODALITY = "SS" Or MODALITY = "BP" Or MODALITY = "HT" Or MODALITY = "ALL" Then
    MSFlexGrid1.FormatString = " Sr.No |^   Requisition No.   |^   Case No.   |^    Procedure Name          |^  Sample Collected  |^    Sample Received |^    DTM NO."
End If
If MODALITY = "BM" Then
    MSFlexGrid1.FormatString = " Sr.No |^   Requisition No.   |^   Case No.   |^    Procedure Name          |^  Sample Collected  |^    Sample Received |^    BMA No. "
End If
If MODALITY = "HL" Or MODALITY = "BF" Then
    MSFlexGrid1.FormatString = " Sr.No |^   Requisition No.   |^   Case No.   |^    Procedure Name          |^  Sample Collected  |^    Sample Received |^    HEM No. "
End If
End Sub

Private Sub CMB_PRINT1_Click()
If Trim(msk_fromdt.Text) = "__/__/____" And Trim(msk_todt.Text) = "__/__/____" Then
    MsgBox "Enter From and To date", vbInformation
    Exit Sub
End If
var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)

CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DEPTCOLLECTIONN.RPT"
'Cr_reqstat.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & Trim(cbo_modality) & " ' AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
CrystalReport1.SelectionFormula = "{RREQUISITN.EXAMDT} >= Date(" & var_date_f & ") AND {RREQUISITN.EXAMDT} <= Date(" & var_date_t & ") and {RREQUISITN.DEPT_CODE}= '" & dept & " '  AND {RREQUISITN.MODaLITYCD}= '" & MODALITY & "'"
CrystalReport1.Formulas(0) = "FROMDT = '" & msk_fromdt.FormattedText & "'"
CrystalReport1.Formulas(1) = "TODT = '" & msk_todt.FormattedText & "'"
CrystalReport1.Formulas(2) = "MODALITY = '" & Left(cmb_modality.Text, 30) & "'"
CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport1.Action = 1
End Sub

Private Sub cmd_back_Click()
chk_rcvd.Value = 0
chk_rcvd.Visible = False
MSFlexGrid1.Rows = 1
BARCODE_NEW = False
BARCODE_VIEW = False
'chk_collected.Value = 0
'chk_rcvd.Value = 0
'chk_collected.Visible = False
'chk_rcvd.Visible = False
txt_pathno.Visible = False
msk_caseno.Mask = ""
msk_caseno.Text = ""
msk_caseno.Mask = "??>/#####"
msk_caseno.Enabled = False
'Msk_fromdt.Mask = ""
'Msk_fromdt.Text = ""
'Msk_fromdt.Mask = "##/##/####"
'Msk_fromdt.Enabled = False
'Msk_todt.Mask = ""
'Msk_todt.Text = ""
'Msk_todt.Mask = "##/##/####"
'Msk_todt.Enabled = False
cmb_dept.ListIndex = -1
cmb_dept.Enabled = False
cmb_modality.ListIndex = -1
cmb_modality.Enabled = False
'opt_patient.Value = False
'opt_staff.Value = False
'Opt_all.Value = False
'opt_caseno.Value = False
btnenabl ("INI")
Me.MousePointer = 11
DIS_MAIN.show
Me.Hide
Me.MousePointer = 0
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
    ' codded for modify option
    If Trim(dept) = "GZZ" Or Trim(dept) = "FEC" Then
        cmd_modify.Enabled = True
    Else
        cmd_modify.Enabled = False
    End If
    ' codded for modify option
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

Private Sub cmd_cancel_Click()
NEWFLAG = False
MSFlexGrid1.Rows = 1
'chk_rcvd.Value = 0
BARCODE_NEW = False
BARCODE_VIEW = False
chk_rcvd.Visible = False
cmb_dept.Text = ""
cmb_modality.Text = ""
'Msk_fromdt.Mask = ""
'Msk_fromdt.Text = ""
'TXT_NAME.Text = ""
'Msk_fromdt.Mask = "##/##/####"
'Msk_todt.Mask = ""
'Msk_todt.Text = ""
'Msk_todt.Mask = "##/##/####"
'Opt_all.Value = False
'Opt_all.Enabled = False
'opt_caseno.Value = False
'opt_caseno.Enabled = False
'opt_staff.Value = False
'opt_staff.Enabled = False
'opt_patient.Value = False
'opt_patient.Enabled = False
msk_caseno.Mask = ""
msk_caseno.Text = ""
msk_caseno.Mask = "??>/#####"
txt_pathno.Text = ""
msk_caseno.Enabled = False
txt_pathno.Visible = False
MODIFYFLAG = False
'chk_collected.Visible = False
btnenabl ("INI")
FLDENABL ("DIS")
End Sub

Private Sub cmd_modify_Click()
btnenabl ("NEW")
'chk_rcvd.Visible = True
txt_pathno.Visible = True
txt_pathno.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6)
'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = ""
MODIFYFLAG = True
' codded for modify option
If dept = "GZZ" Then
    If viewflag = True Then
        MSFlexGrid1_Click
        txt_pathno.Visible = True
    End If
End If
' codded for modify option
End Sub

Private Sub cmd_new_Click()
'On Error GoTo MSG:
chk_rcvd.Value = 0
BARCODE_NEW = True
If SELMA106FLAG = False Then
    Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
    SELMA106FLAG = True
Else
    Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
    SELMA106_RS.MoreResults
    SELMA106_RS.Requery
End If

If SELMA106_RS.RowCount > 0 Then
    cmb_dept.clear
    While Not SELMA106_RS.EOF
        If Trim(USRDEPT) <> "ALL" Then
            If Trim(USRDEPT) = Trim(SELMA106_RS("DEPT_CODE")) Then
                cmb_dept.AddItem Trim(SELMA106_RS("DEPT_NAME")) & "$" & Trim(SELMA106_RS("DEPT_CODE"))
            End If
        Else
             cmb_dept.AddItem Trim(SELMA106_RS("DEPT_NAME")) & "$" & Trim(SELMA106_RS("DEPT_CODE"))
        End If
        SELMA106_RS.MoveNext
    Wend
End If

cmb_dept.Text = cmb_dept.List(0)

If Trim(cmb_dept.Text) = Trim("PATHOLOGY$FZZ") Then
    cmb_dept.AddItem Trim("PREVENTIVE ONCOLOGY$IZZ")
End If
If Trim(cmb_dept.Text) <> "" Then
    cmb_dept_Click
End If
MODIFYFLAG = False
'Opt_all.Enabled = True
'opt_caseno.Enabled = True
'opt_patient.Enabled = True
'opt_staff.Enabled = True
txt_name.Text = ""
chk_rcvd.Visible = False
btnenabl ("NEW")
FLDENABL ("ENA")
NEWFLAG = True
viewflag = False
MSFlexGrid1.Rows = 2
MSFlexGrid1.row = 1

'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
'opt_patient.Value = True
cmb_dept.SetFocus

Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_print_Click()
Dim var_date_f
Dim var_date_t
Dim str1 As String
Dim str As String
str = "C"
On Error GoTo MSG:
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    If OPT_ALL.Value = True Then
        var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
        var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
        If Right(Trim(cmb_modality.Text), 2) = "SS" Or Right(Trim(cmb_modality.Text), 2) = "BP" Or Right(Trim(cmb_modality.Text), 2) = "HT" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DEPTCOLLECTIONDTM.RPT"
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DEPTCOLLECTION.RPT"
        End If
        'Cr_reqstat.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & Trim(cbo_modality) & " ' AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
        CrystalReport1.SelectionFormula = "{RREQUISITN.EXAMDT} >= Date(" & var_date_f & ") AND {RREQUISITN.EXAMDT} <= Date(" & var_date_t & ") and {RREQUISITN.DEPT_CODE}= '" & dept & " '  AND {RREQUISITN.MODALITYCD}= '" & MODALITY & "' "
        CrystalReport1.Formulas(0) = "FROMDT = '" & msk_fromdt.FormattedText & "'"
        CrystalReport1.Formulas(1) = "TODT = '" & msk_todt.FormattedText & "'"
        CrystalReport1.Formulas(2) = "MODALITY = '" & Left(cmb_modality.Text, 30) & "'"
        CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
    End If

    If OPT_CASENO.Value = True Then
        If Right(Trim(cmb_modality.Text), 2) = "SS" Or Right(Trim(cmb_modality.Text), 2) = "HT" Or Right(Trim(cmb_modality.Text), 2) = "BP" Then
            str1 = msk_caseno.FormattedText
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DEPTCOLLECTIONCDTM.RPT"
            'CRYSTALREPORT1.SelectionFormula = " {RREQUISITN.CASENO} ='" & msk_caseno.FormattedText & "' and {RREQUISITN.DEPT_CODE}= '" & DEPT & " '  AND {RREQUISITN.MODaLITYCD}= '" & MODALITY & "' AND ({RREQUISITN.REQNSTAT}= 'E' OR {RREQUISITN.REQNSTAT}= 'C') "
            CrystalReport1.SelectionFormula = " {RREQUISITN.CASENO} ='" & msk_caseno.FormattedText & "' and {RREQUISITN.DEPT_CODE}= '" & dept & " '  AND {RREQUISITN.MODaLITYCD}= '" & MODALITY & "' AND ({RREQUISITN.REQNSTAT}= 'E') "
            CrystalReport1.Formulas(0) = "FROMDT = '" & msk_caseno.FormattedText & "'"
            CrystalReport1.Formulas(1) = "MODALITY = '" & Left(cmb_modality.Text, 30) & "'"
        Else
            str1 = msk_caseno.FormattedText
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DEPTCOLLECTIONC.RPT"
            CrystalReport1.SelectionFormula = " {RREQUISITN.CASENO} ='" & msk_caseno.FormattedText & "' and {RREQUISITN.DEPT_CODE}= '" & dept & " '  AND {RREQUISITN.MODaLITYCD}= '" & MODALITY & "' AND ({RREQUISITN.REQNSTAT}= 'E' OR {RREQUISITN.REQNSTAT}= 'C')"
            CrystalReport1.Formulas(0) = "FROMDT = '" & msk_caseno.FormattedText & "'"
            CrystalReport1.Formulas(1) = "MODALITY = '" & Left(cmb_modality.Text, 30) & "'"
        End If
    End If

    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

        CrystalReport1.Action = 1

btnenabl ("PRINT")
Exit Sub
MSG:
    MsgBox Err.Description

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
txt_pathno.Text = ""
chk_rcvd.Visible = False
txt_pathno.Visible = False
btnenabl ("NEW")
MODIFYFLAG = False
'chk_rcvd.Visible = False
If msk_caseno.Enabled = True Then
    msk_caseno.SetFocus
End If
End Sub

Private Sub cmd_save_Click()
Dim ST As String
datetime
Dim hlreqno As String
Dim path_nosp As String
Dim Case_nosp As String
Dim Locationsp As String
txt_time.Text = time1
'On Error GoTo MSG:
Dim REC As Double
ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
If MODALITY = "SP" Or MODALITY = "HL" Or MODALITY = "BM" Or MODALITY = "BF" Then
    For i = 1 To MSFlexGrid1.Rows - 1
        If MSFlexGrid1.TextMatrix(i, 5) <> "" Then
            REC = REC + 1
        End If
    Next i
    If MsgBox("U'r Going to Save " & Trim(REC) & " Records, is it OK ?", vbOKCancel) = vbOK Then
    'If MODALITY = "SP" Then
        SPAUTONO
    'Else
    'HEMAUTONO
    'End If
    Else
        Exit Sub
    End If
End If

'If MODALITY = "TM" Then
'   TMAUTONO
'End If
'For i = 1 To MSFlexGrid1.Rows - 1
'    If MSFlexGrid1.TextMatrix(i, 6) <> "" Then
'        If modality = "BC" Or modality = "EL" Then
'            If seldequisitnflag = False Then
'                seldrequisitnqr(0) = st & UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
'                Set seldrequisitnrs = seldrequisitnqr.OpenResultset(1, 3)
'                seldequisitnflag = True
'            Else
'                seldrequisitnqr(0) = st & UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
'                seldrequisitnrs.MoreResults
'                seldrequisitnrs.Requery
'            End If
'        Else
'            If seldequisitnflag = False Then
'                seldrequisitnqr(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
'                Set seldrequisitnrs = seldrequisitnqr.OpenResultset(1, 3)
'                seldequisitnflag = True
'            Else
'                seldrequisitnqr(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
'                seldrequisitnrs.MoreResults
'                seldrequisitnrs.Requery
'            End If
'        End If
'        If MODIFYFLAG = False Then
'            If seldrequisitnrs.RowCount > 0 Then
'                MsgBox "The Path No. is  Given already " & seldrequisitnrs("caseno")
'                txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
'                'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
'                txt_pathno.Visible = True
'                'chk_rcvd.Visible = True
'                'chk_rcvd.SetFocus
'                txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'                'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
'                Exit Sub
'
'            End If
'        Else
'            MODIFYFLAG = True
'        End If
'    End If
'Next i

rdb.BeginTrans
If NEWFLAG = True Then
    For i = 1 To MSFlexGrid1.Rows - 1
        'MSFlexGrid1.Col = 4
        srow = i
        If Trim(dept) = "NZZ" Then
            If Trim(MSFlexGrid1.TextMatrix(i, 4)) <> "" And Trim(MSFlexGrid1.TextMatrix(i, 5)) <> "" Then
                updrrequisitn_qr(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))

                If SELREQNORSFLAG = False Then
                    SELREQNOQR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    Set SELREQNORS = SELREQNOQR.OpenResultset(1, 3)
                    SELREQNORSFLAG = True
                Else
                    SELREQNOQR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    SELREQNORS.MoreResults
                    SELREQNORS.Requery
                End If

                If Trim(SELREQNORS("reqnstat")) = "P" Then
                    updrrequisitn_qr(1) = "E"
                End If
                If Trim(SELREQNORS("reqnstat")) = "C" Then
                    updrrequisitn_qr(1) = "E"
                End If
                If Trim(SELREQNORS("REQNSTAT")) = "E" Then
                    updrrequisitn_qr(1) = "E"
                End If
                If Trim(SELREQNORS("REQNSTAT")) = "T" Then
                    updrrequisitn_qr(1) = "T"
                End If
                If Trim(SELREQNORS("REQNSTAT")) = "F" Then
                    updrrequisitn_qr(1) = "F"
                End If
                updrrequisitn_qr(2) = "RECEIVED"
                If Left(Trim(MSFlexGrid1.TextMatrix(i, 2)), 2) = "RF" Or Left(Trim(MSFlexGrid1.TextMatrix(i, 2)), 2) = "RG" Then
                    updrrequisitn_qr(3) = "COLLRCVD"
                    'Locationsp = "COLLRCVD"
                Else
                    If Len(Trim(SELREQNORS("LOCATION"))) = "" Then
                        updrrequisitn_qr(3) = "COLLECTED"
                       ' Locationsp = "COLLECTED"
                    Else
                       updrrequisitn_qr(3) = Trim(SELREQNORS("LOCATION"))
                      ' Locationsp = Trim(SELREQNORS("LOCATION"))
                    End If
                End If
                updrrequisitn_qr(4) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                updrrequisitn_qr(5) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                updrrequisitn_qr.Execute

                ''RECEIPT TIME

                UPDLATERALITS_QR(0) = time1
                UPDLATERALITS_QR(1) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                UPDLATERALITS_QR.Execute
                ''***********

                If dept <> "NZZ" Then
                    If MODALITY = "BC" Or MODALITY = "EL" Or MODALITY = "TM" Then
                        If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) > 3 Then
                            upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                        End If
                        If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <> 0 Then
                            upddrequisitn(0) = ST & UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                        End If
                        '            If modality = "BC" Then
                        '                BCINF
                        '            End If
                        '            If modality = "EL" Then
                        '                ELINF
                        '            End If
                    Else
                        upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                    End If
                    upddrequisitn(1) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    upddrequisitn.Execute
                End If

                UPDTIME_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
                UPDTIME_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                UPDTIME_QR(2) = msk_caseno.FormattedText
                UPDTIME_QR(3) = SELREQNORS("REQNDT")
                UPDTIME_QR(4) = SELREQNORS("REQNTM")
                UPDTIME_QR(5) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                UPDTIME_QR.Execute

            End If

        Else
            If MSFlexGrid1.TextMatrix(i, 4) <> "" And MSFlexGrid1.TextMatrix(i, 5) <> "" And MSFlexGrid1.TextMatrix(i, 6) <> "" Then
                updrrequisitn_qr(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                If SELREQNORSFLAG = False Then
                    SELREQNOQR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    'FETCH REQNO  IN BARCODE AMIT
                    REQNO = SELREQNOQR(0)
                    Set SELREQNORS = SELREQNOQR.OpenResultset(1, 3)
                    SELREQNORSFLAG = True
                Else
                    SELREQNOQR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    SELREQNORS.MoreResults
                    SELREQNORS.Requery
                End If
                If Trim(SELREQNORS("reqnstat")) = "P" Then
                    updrrequisitn_qr(1) = "E"
                End If
                If Trim(SELREQNORS("reqnstat")) = "C" Then
                    updrrequisitn_qr(1) = "E"
                End If
                If Trim(SELREQNORS("REQNSTAT")) = "E" Then
                    updrrequisitn_qr(1) = "E"
                End If
                If Trim(SELREQNORS("REQNSTAT")) = "T" Then
                    updrrequisitn_qr(1) = "T"
                End If
                If Trim(SELREQNORS("REQNSTAT")) = "F" Then
                    updrrequisitn_qr(1) = "F"
                End If
                updrrequisitn_qr(2) = "RECEIVED"
                If Left(Trim(MSFlexGrid1.TextMatrix(i, 2)), 2) = "RF" Or Left(Trim(MSFlexGrid1.TextMatrix(i, 2)), 2) = "RG" Then
                    updrrequisitn_qr(3) = "COLLRCVD"
                    Locationsp = "COLLRCVD"
                Else
                    If Len(Trim(SELREQNORS("LOCATION"))) = 0 Then  ' made 0 from ""
                        updrrequisitn_qr(3) = "COLLECTED"
                        Locationsp = "COLLECTED"
                    Else
                       updrrequisitn_qr(3) = Trim(SELREQNORS("LOCATION"))
                       If IsNull(Trim(SELREQNORS("LOCATION"))) = True Then
                        Locationsp = ""
                       Else
                        Locationsp = Trim(SELREQNORS("LOCATION"))
                       End If
                    End If
                End If
                updrrequisitn_qr(4) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                updrrequisitn_qr(5) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                updrrequisitn_qr.Execute

                UPDTIME_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
                UPDTIME_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                UPDTIME_QR(2) = msk_caseno.FormattedText
                UPDTIME_QR(3) = SELREQNORS("REQNDT")
                UPDTIME_QR(4) = SELREQNORS("REQNTM")
                UPDTIME_QR(5) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                UPDTIME_QR.Execute

                SELTIME_QR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                If SELTIME_FLG = False Then
                    Set SELTIME_RS = SELTIME_QR.OpenResultset(1, 3)
                    SELTIME_FLG = True
                Else
                  SELTIME_RS.MoreResults
                  SELTIME_RS.Requery
                End If
                'If SELTIME_RS.RowCount = 0 Then
                ''THIS IS FOR INSERT IN RTIMELOG
                'INSTIME_QR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                'INSTIME_QR(1) = msk_caseno.FormattedText
                'INSTIME_QR(2) = SELREQNORS("REQNDT")
                'INSTIME_QR(3) = SELREQNORS("REQNTM")
                'INSTIME_QR(4) = Format(txt_date.Text, "yyyy-mm-dd")
                'INSTIME_QR(5) = Format(txt_time.Text, "HH:MM:SS")
                'INSTIME_QR.Execute
                '   INSTIME_QR.Close
                'Else
                '   UPDTIME_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
                '   UPDTIME_QR(1) = Format(txt_time.Text, "HH:MM:SS")
                '   UPDTIME_QR(2) = msk_caseno.FormattedText
                '   UPDTIME_QR(3) = SELREQNORS("REQNDT")
                '   UPDTIME_QR(4) = SELREQNORS("REQNTM")
                '   UPDTIME_QR(5) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                '   UPDTIME_QR.Execute
                'End If
                If MODALITY = "BC" Or MODALITY = "EL" Or MODALITY = "TM" Then
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) > 3 Then
                        upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                    End If
                    If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <> 0 Then
                        upddrequisitn(0) = ST & UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                    End If
                    'If modality = "BC" Then
                    '   BCINF
                    'End If
                    'If modality = "EL" Then
                    '   ELINF
                    'End If
                Else
                    'COMMENTED BY MILIND ON 31-03-2007 TO DISABLE GENERATING
                    'PATH NO FOR SAME REQUISITION NO.
                    l = UCase(MSFlexGrid1.TextMatrix(i, 6))
                    'upddrequisitn(0) = l
                    upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                    path_nosp = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                End If
                'upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
                upddrequisitn(1) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                upddrequisitn.Execute

                ' Insert value into SPathology from here
                If MODALITY = "SP" Then
                'path_no,case_no,trans_dt,loct1
                    Ins_Spathology_qr(0) = path_nosp
                    Ins_Spathology_qr(1) = MSFlexGrid1.TextMatrix(i, 2)
                    Ins_Spathology_qr(2) = Format(DATE11, "YYYY/MM/DD")
                    Ins_Spathology_qr(3) = Locationsp
                    Ins_Spathology_qr.Execute
                End If
                ' Insert value into SPathology from here
                ''  HEMATO INTERFACE CODE BEGINS
                If MODALITY = "HL" Then
                    hlreqno = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    SELDHEMCYTO_QR(0) = hlreqno  ' req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    Set SELDHEMCYTO_RS = SELDHEMCYTO_QR.OpenResultset(1, 1)
                    If SELDHEMCYTO_RS.RowCount > 0 Then
                        If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("HEM")) = "YES" Then
                            INSHINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                            INSHINTERFACE_QR(1) = hlreqno
                            INSHINTERFACE_QR(2) = " "
                            INSHINTERFACE_QR(3) = "1"
                            INSHINTERFACE_QR(4) = Trim(Mid(hlreqno, 4, 10)) + "A"
                            INSHINTERFACE_QR.Execute
                        End If
                        If Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
                            seldmulproce_qr(0) = hlreqno
                            Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 1)
                            Do While seldmulproce_rs.EOF = False
                                seldhlprosub_qr(0) = seldmulproce_rs("pro_code")
                                Set seldhlprosub_rs = seldhlprosub_qr.OpenResultset(1, 1)
                                Do While seldhlprosub_rs.EOF = False
                                    ' STR = "INSERT INTO INTERFACE.CINTERFACE(CASENO,REQNO,TESTCODE,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?,?)"
                                    If seldhlprosub_rs("samplecd") = "B" Then
                                        INSCINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                                        INSCINTERFACE_QR(1) = hlreqno
                                        INSCINTERFACE_QR(2) = seldhlprosub_rs("procodesub")
                                        INSCINTERFACE_QR(3) = " "
                                        INSCINTERFACE_QR(4) = "1"
                                        INSCINTERFACE_QR(5) = Trim(Mid(hlreqno, 4, 10)) + Trim(seldhlprosub_rs("samplecd"))
                                        INSCINTERFACE_QR.Execute
                                    End If
                                    seldhlprosub_rs.MoveNext
                                Loop
                                seldmulproce_rs.MoveNext
                            Loop
                        End If
                    End If 'dhemctyo rowcount
                End If  'HL

                'HEMATO INTERFACE ENDS

                If MODALITY = "BC" Then
                    If SELDMULPROCE_FLAG = False Then
                      seldmulproce_qr(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                      Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 3)
                      SELDMULPROCE_FLAG = True
                    Else
                      seldmulproce_qr(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                      seldmulproce_rs.MoreResults
                      seldmulproce_rs.Requery
                    End If
                    REQNO = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    If Not seldmulproce_rs.EOF Then
                      Do While Not seldmulproce_rs.EOF
                        'checking for record in excption codes which are used as manual
                        ' and not in interface
                        'WRITTE BY AMIT STORING DATA IN INTERFACE/DINTERFACE
                    PCODE = seldmulproce_rs("PRO_CODE")
                    If PCODE = "F849" Or PCODE = "F854" Or PCODE = "F924" Or PCODE = "F860" Or PCODE = "F875" Or PCODE = "F880" Or PCODE = "F875" Or PCODE = "F865" Or PCODE = "F802" Then
                        SELDBCPROSUB_QR(0) = PCODE
                        Set SELDBCPROSUB_RS = SELDBCPROSUB_QR.OpenResultset(1, 3)
                        Do Until SELDBCPROSUB_RS.EOF
                            interexpcd_qr(0) = SELDBCPROSUB_RS!PROCODESUB
                            Set interexpxd_rs = interexpcd_qr.OpenResultset(1, 1)
                            'If interexpxd_rs.RowCount = 0 Then
                                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                                Dim REQ1, P_CODE, SCODE
                                REQ1 = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                                P_CODE = SELDBCPROSUB_RS!PROCODESUB
                                SELSAMPLECODE_QR(0) = P_CODE
                                Set SAMPLECODE_RS = SELSAMPLECODE_QR.OpenResultset(1, 3)
                                SCODE = SAMPLECODE_RS!SAMPLECD
                                REQ2 = Mid$(REQ1, 4)
                                SCODE = REQ2 & SCODE
                                INSDINTERFACE_QR(1) = REQ1
                                INSDINTERFACE_QR(2) = SELDBCPROSUB_RS!PROCODESUB
                                INSDINTERFACE_QR(3) = ""
                                INSDINTERFACE_QR(4) = "1"
                                INSDINTERFACE_QR(5) = SCODE
                                INSDINTERFACE_QR(6) = Format(Date, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(i, 6))
                                INSDINTERFACE_QR.Execute
                            'End If
                            SELDBCPROSUB_RS.MoveNext
                        Loop
                     Else
                        interexpcd_qr(0) = seldmulproce_rs("PRO_CODE")
                        Set interexpxd_rs = interexpcd_qr.OpenResultset(1, 1)
                        If interexpxd_rs.RowCount = 0 Then
                            INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                                Dim REQ11, P_CODE1, SCODE1
                                REQ11 = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                                SELSAMPLECODE_QR(0) = PCODE
                                Set SAMPLECODE_RS = SELSAMPLECODE_QR.OpenResultset(1, 3)
                                SCODE1 = SAMPLECODE_RS!SAMPLECD
                                REQ2 = Mid$(REQ11, 4)
                                SCODE1 = REQ2 & SCODE1

                            INSDINTERFACE_QR(1) = REQ11
                            INSDINTERFACE_QR(2) = seldmulproce_rs("PRO_CODE")
                            INSDINTERFACE_QR(3) = ""
                            INSDINTERFACE_QR(4) = "1"
                            INSDINTERFACE_QR(5) = SCODE1
                            INSDINTERFACE_QR(6) = Format(Date, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(i, 6))
                            INSDINTERFACE_QR.Execute
                        End If
                    End If
                    seldmulproce_rs.MoveNext
                    Loop
                    End If

                ' END BY AMIT

                  End If


                  'TM INTERFACE 2005

                  If MODALITY = "TM" Then
                    If SELDMULPROCE_FLAG = False Then
                      seldmulproce_qr(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                      Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 3)
                      SELDMULPROCE_FLAG = True
                    Else
                      seldmulproce_qr(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                      seldmulproce_rs.MoreResults
                      seldmulproce_rs.Requery
                    End If
                    REQNO = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                    If Not seldmulproce_rs.EOF Then
                      Do While Not seldmulproce_rs.EOF
                         PCODE = seldmulproce_rs("PRO_CODE")
                        interexpcd_qr(0) = seldmulproce_rs("PRO_CODE")
                        Set interexpxd_rs = interexpcd_qr.OpenResultset(1, 1)
                        If interexpxd_rs.RowCount = 0 Then
                            INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                                Dim REQ11t, P_CODE1t, SCODE1t
                                REQ11t = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
                                SELSAMPLECODE_QR(0) = PCODE
                                Set SAMPLECODE_RS = SELSAMPLECODE_QR.OpenResultset(1, 3)
                                If SAMPLECODE_RS.RowCount > 0 Then

                                 SCODE1t = SAMPLECODE_RS!SAMPLECD
                                Else
                                  SCODE1t = " "
                                End If

                                REQ2 = Mid$(REQ11t, 4)
                                SCODE1t = REQ2 & SCODE1t

                            INSDINTERFACE_QR(1) = REQ11t
                            INSDINTERFACE_QR(2) = seldmulproce_rs("PRO_CODE")
                            INSDINTERFACE_QR(3) = ""
                            INSDINTERFACE_QR(4) = "1"
                            INSDINTERFACE_QR(5) = SCODE1t
                            INSDINTERFACE_QR(6) = Format(Date, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(i, 6))
                            INSDINTERFACE_QR.Execute
                        End If
                     seldmulproce_rs.MoveNext
                     Loop
                   End If
                   '' FOR SAVING BINTERFACE
                   Set RS_DMULPROCE = rdb.OpenResultset("SELECT * FROM DISLIB.DMULPROCE WHERE REQNO='" & REQNO & "' AND PRO_CODE='F816' ", 1, 1)
                      If RS_DMULPROCE.RowCount > 0 Then
                        Dim req As Variant
                        req = Mid$(REQNO, 4)
                        INSBINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(i, 2)
                        INSBINTERFACE_QR(1) = REQNO
                        INSBINTERFACE_QR(2) = ""
                        INSBINTERFACE_QR(3) = "1"
                        INSBINTERFACE_QR(4) = req & "A"
                        INSBINTERFACE_QR.Execute
                      End If
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''1

                 End If

                'TM INTERFACE 2005
              End If

            End If

          Next i

      MsgBox "Record is Successfully Saved."

End If
If MODIFYFLAG = True Then

For i = 1 To MSFlexGrid1.Rows - 1
     'MSFlexGrid1.Col = 4
    If MSFlexGrid1.TextMatrix(i, 4) <> "" And MSFlexGrid1.TextMatrix(i, 5) <> "" And MSFlexGrid1.TextMatrix(i, 6) <> "" Then
        updrrequisitn_qr(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        If SELREQNORSFLAG = False Then
            SELREQNOQR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
            Set SELREQNORS = SELREQNOQR.OpenResultset(1, 3)
            SELREQNORSFLAG = True
        Else
            SELREQNOQR(0) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
            SELREQNORS.MoreResults
            SELREQNORS.Requery
        End If
        If Trim(SELREQNORS("reqnstat")) = "P" Then
            updrrequisitn_qr(1) = "E"
        End If
        If Trim(SELREQNORS("reqnstat")) = "C" Then
            updrrequisitn_qr(1) = "E"
        End If
        If Trim(SELREQNORS("REQNSTAT")) = "E" Then
            updrrequisitn_qr(1) = "E"
        End If
        If Trim(SELREQNORS("REQNSTAT")) = "T" Then
            updrrequisitn_qr(1) = "T"
        End If
        If Trim(SELREQNORS("REQNSTAT")) = "F" Then
            updrrequisitn_qr(1) = "F"
        End If
        updrrequisitn_qr(2) = "RECEIVED"
        If Left(Trim(MSFlexGrid1.TextMatrix(i, 2)), 2) = "RF" Or Left(Trim(MSFlexGrid1.TextMatrix(i, 2)), 2) = "RG" Then
            updrrequisitn_qr(3) = "COLLRCVD"
        Else
            updrrequisitn_qr(3) = "COLLECTED"
        End If
        updrrequisitn_qr(4) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        updrrequisitn_qr(5) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
        updrrequisitn_qr.Execute
        If MODALITY = "BC" Or MODALITY = "EL" Or MODALITY = "TM" Then
            If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) > 3 Then
            upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
            End If
            If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <> 0 Then
             upddrequisitn(0) = ST & UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
            End If
            If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) = 0 Then
             upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
            End If
        Else
            upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
        End If
        upddrequisitn(1) = req_chglen(MSFlexGrid1.TextMatrix(i, 1))
        upddrequisitn.Execute
    End If
Next i
MsgBox "Record is Successfully Saved."
End If
rdb.CommitTrans

btnenabl ("SAVE")
Exit Sub
'MSG:
rdb.RollbackTrans
'MsgBox Err.Description
btnenabl ("SAVE")
End Sub

Private Sub cmd_view_Click()
'On Error GoTo MSG:
BARCODE_VIEW = True
chk_rcvd.Value = 0
viewflag = True
NEWFLAG = False
If SELMA106FLAG = False Then
    Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
    SELMA106FLAG = True
Else
    Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
    SELMA106_RS.MoreResults
    SELMA106_RS.Requery

End If
If SELMA106_RS.RowCount > 0 Then
    cmb_dept.clear
    While Not SELMA106_RS.EOF
        cmb_dept.AddItem Trim(SELMA106_RS("DEPT_NAME")) & "$" & Trim(SELMA106_RS("DEPT_CODE"))
        SELMA106_RS.MoveNext
    Wend
End If
txt_name.Text = ""
'Opt_all.Enabled = True
'Opt_all.Value = False
'opt_caseno.Enabled = True
'Opt_all.Value = False
'opt_patient.Enabled = True
'opt_patient.Value = False
'opt_staff.Enabled = True
'opt_staff.Value = False
chk_rcvd.Value = 0
chk_rcvd.Visible = False

btnenabl ("VIEW")
FLDENABL ("ENA")
cmb_dept.SetFocus
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub cmd_VPP_Click()
On Error GoTo MSG:
CrystalReport1.ReportFileName = "C:\DISREPORTSEC\vpp.RPT"
CrystalReport1.SelectionFormula = " {RREQUISITN.REQNNO} ='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1) & "'"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CrystalReport1.Action = 1
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
'On Error GoTo MSG:
Main
If deptcollectionflag = False Then
datetime
txt_date.Text = Format(DATE11, "dd/mm/yyyy")
txt_time.Text = time1
'chk_rcvd.Visible = False
txt_pathno.Visible = False
deptcollectionflag = True
scrlck_flg = True
MODIFYFLAG = False
'INTERFACE
str = "SELECT DISTINCT LABNO FROM INTERFACE.DINTERFACE WHERE REQNO=?"
Set SELLABNO_QR = rdb.CreateQuery("LABNO", str)

str = "SELECT PATHNO FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELPATH_QR = rdb.CreateQuery("PATHNO", str)

str = "SELECT TESTCODE FROM INTERFACE.DINTERFACE WHERE REQNO=?"
Set SELPROCODE_QR = rdb.CreateQuery("PROCODE", str)

' FOR RTIMELOG
'STR = "INSERT INTO RISLIB.RTIMELOG(REQNNO,CASENO,REQNDATE,REQNTIME) VALUES(?,?,?,?)"
'STR = "INSERT INTO RISLIB.RTIMELOG(REQNNO,CASENO,REQNDATE,REQNTIME,SAMRECDATE,SAMRECTIME) VALUES(?,?,?,?,?,?)"
'Set INSTIME_QR = rdb.CreateQuery("FHGSDF", STR)

''***************INTERFACE**********************
str = "INSERT INTO INTERFACE.DINTERFACE(CASENO,REQNO,TESTCODE,URGENT,IFLAG,LABNO,FLD2) VALUES(?,?,?,?,?,?,?)"
Set INSDINTERFACE_QR = rdb.CreateQuery("INSDINTERFACE_STR", str)

str = "INSERT INTO INTERFACE.BINTERFACE(CASENO,REQNO,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?)"
Set INSBINTERFACE_QR = rdb.CreateQuery("INSBINTERFACE_STR", str)


'' HEMATO INTERFACE
str = "INSERT INTO INTERFACE.HINTERFACE(CASENO,REQNO,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?)"
Set INSHINTERFACE_QR = rdb.CreateQuery("INSHINTERFACE_STR", str)
str = "INSERT INTO INTERFACE.CINTERFACE(CASENO,REQNO,TESTCODE,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?,?)"
Set INSCINTERFACE_QR = rdb.CreateQuery("INSCINTERFACE_STR", str)
str = "SELECT * FROM DISLIB.DHEMCYTO1 WHERE REQNO = ?"
Set SELDHEMCYTO_QR = rdb.CreateQuery("SELDHEMCYTO1_QR", str)

'STR = "select * from dislib.dmulproce where reqno = ?"
str = "select * from dislib.dhlprosub where hlprocode = ?"
Set seldhlprosub_qr = rdb.CreateQuery("SELDHLPROSUB_QR1", str)

'HEMATO INTERFACE

str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
Set seldmulproce_qr = rdb.CreateQuery("SELDMULPROCE_STRDD", str)

str = "SELECT * FROM DISLIB.INTEREXPCD WHERE excpcode =?"
Set interexpcd_qr = rdb.CreateQuery("interexpcd_qr1", str)

str = "SELECT * FROM DISLIB.DBCPROSUB WHERE BCPROCODE=?"
Set SELDBCPROSUB_QR = rdb.CreateQuery("DBCPROSUB_SEL", str)

str = "SELECT SAMPLECD FROM DISLIB.SAMPLECODE WHERE SERCODE=?"
Set SELSAMPLECODE_QR = rdb.CreateQuery("SAMPLE1", str)
''***************INTERFACE**********************
str = "SELECT SAMRECDATE,SAMRECTIME FROM RISLIB.RTIMELOG WHERE REQNNO=?"
Set SELTIME_QR = rdb.CreateQuery("FFHSHF", str)

str = "UPDATE RISLIB.RTIMELOG SET SAMRECDATE=?,SAMRECTIME=?,CASENO=?,REQNDATE=?, REQNTIME=? WHERE REQNNO=? "
Set UPDTIME_QR = rdb.CreateQuery("DFHWH", str)
''*********tracking System*************
str = "insert into dislib.spathology(path_no,case_no,trans_dt,loct1) values(?,?,?,?)"
Set Ins_Spathology_qr = rdb.CreateQuery("Ins_Spathology_qr", str)
''*********tracking System*************

str = "INSERT INTO DISLIB.DAUTOPATH (DEPT_CODE,MODALITYCD,PATHYEAR,PATHYEAR1,PATHNO) VALUES (?,?,?,?,?)"
Set INSDAUTOPATH_QR = rdb.CreateQuery("INSDAUTOPATH_QRXS1", str)

str = "UPDATE DISLIB.DAUTOPATH SET PATHNO=? WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
Set UPDDAUTOPATH_QR = rdb.CreateQuery("UPDDAUTOPATH_Q11R1", str)

' ***** AUTOGENERATION OF THE PATHNO.
str = "SELECT * FROM DISLIB.DAUTOPATH WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
Set SELDAUTOPATH_QR = rdb.CreateQuery("SEXXLDAUTOPATH_QR1", str)

' ***** AUTOGENERATION OF THE PATHNO(TUMOUR MARKER).
str = "SELECT * FROM DISLIB.AUTOPATHTM WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR=? AND PATHMONTH=? AND PATHDATE=?"
Set SELTMAUTOPATH_QR = rdb.CreateQuery("SEXXLTMAUTOPATH_QR1", str)

str = "INSERT INTO DISLIB.AUTOPATHTM (DEPT_CODE,MODALITYCD,PATHYEAR,PATHMONTH,PATHDATE,PATHNO) VALUES (?,?,?,?,?,?)"
Set INSTMAUTOPATH_QR = rdb.CreateQuery("INSTMAUTOPATH_QRXS1", str)

str = "UPDATE DISLIB.AUTOPATHTM SET PATHNO=? WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR=? AND PATHMONTH=? AND PATHDATE=?"
Set UPDTMAUTOPATH_QR = rdb.CreateQuery("UPDTMAUTOPATH_Q11R1", str)


str = "SELECT * FROM DISLIB.DPATHYEAR WHERE AYEAR=?"
Set DPATHYEAR_QR = rdb.CreateQuery("DPXXATHYEAR_QR1", str)

''TO AVOID DUPLICATE REQUISITION RECEIVING
str = "SELECT PATHNO FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELPATHNO_QR = rdb.CreateQuery("SELPATHNO_QR", str)
''''

'''RECEIPT TIME
str = "UPDATE DISLIB.DREQUISITN SET LATERALITS=? WHERE REQNO=?"
Set UPDLATERALITS_QR = rdb.CreateQuery("UPDLATERALITS_STR", str)

''DTMNO
str = "SELECT * FROM DISLIB.DTMNO WHERE CASENO=?"
Set SELDTMNO_QR = rdb.CreateQuery("SELDTMNO_STR11", str)

str = "select * from patoolslib.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ'  OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ' or dept_code='IZZ')"
Set SELMA106_QR = rdb.CreateQuery("selma106_qr11", str)

str = "select * from Rislib.rrequisitn where dept_code=? and modalitycd=? and (reqndt between ? and ?) AND (REQNSTAT=? or REQNSTAT=?) ORDER BY CASENO "
Set SELRREQUISITN_QR = rdb.CreateQuery("selrrequisitn_qr111", str)

str = "select * from Rislib.rrequisitn where dept_code='NZZ' and (reqndt between ? and ?) AND (REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=?) ORDER BY CASENO "
Set SELRREQUISITN_QR1 = rdb.CreateQuery("selrrequisitn_qr11", str)

str = "select * from Rislib.rrequisitn where dept_code='NZZ' and (reqndt between ? and ?) AND MODALITYCD=? AND (REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=?) ORDER BY CASENO "
Set SELRREQUISITN_QR11 = rdb.CreateQuery("selrrequisitn_qr10091", str)


str = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO_STR111", str)

'STR = "SELECT * FROM RISLIB.RREQUISITN WHERE DEPT_CODE=? AND MODALITYCD=? AND CASENO=? AND (REQNSTAT=? OR REQNSTAT=?)"
str = "SELECT * FROM RISLIB.RREQUISITN WHERE DEPT_CODE=? AND MODALITYCD=? AND CASENO=? AND (REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=? )"
Set SELRREQNO_QR = rdb.CreateQuery("SELRREQNO1_QR1", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE DEPT_CODE='NZZ' AND CASENO=? AND (REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=? )"
Set SELRREQNO_QR1 = rdb.CreateQuery("SELRREQNO1_QR1001", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE DEPT_CODE='NZZ' AND MODALITYCD=? AND CASENO=? AND (REQNSTAT=? OR REQNSTAT=? OR REQNSTAT=?)"
Set SELRREQNO_QR11 = rdb.CreateQuery("CCVSELRREQNO1_QR1001", str)

str = "select * from rislib.rmodality where dept_code=?"
Set SELRMODALITY_QR = rdb.CreateQuery("selrmodality_qr1", str)

str = "select * from patoolslib.ma101 where ser_code=?"
Set SELMA101_QR = rdb.CreateQuery("selma1011_qr1", str)

str = "UPDATE  RISLIB.RREQUISITN SET EXAMDT=?, REQNSTAT=?,EXAMINATN=?,LOCATION=?,APNTDT=? WHERE REQNNO=?"
Set updrrequisitn_qr = rdb.CreateQuery("UPDRREQUISITN_QR11", str)

str = "select * from patoolslib.ta101 where case_no=?"
Set SELTA1011_QR = rdb.CreateQuery("selta10111_qr1", str)

str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE EMPCODE=?"
Set SELDREQSTAFFD1_QR = rdb.CreateQuery("SELDREQSTAFFD_QR11", str)

str = "SELECT * FROM DISLIB.DREQUISITN WHERE PATHNO=? AND  REQNO LIKE ?"
'Commented and added by Shraddha on 19/07/10
'str = "SELECT * FROM DISLIB.DREQLF_R WHERE PATHNO=? AND  REQNO LIKE ?"
Set SELDREQUISITNQR = rdb.CreateQuery("1SELDREQUISITNQR1", str)

str = "UPDATE DISLIB.DREQUISITN SET PATHNO=? WHERE REQNO=?"
Set upddrequisitn = rdb.CreateQuery("UPDRREQUISITN1", str)

str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELDREQQR = rdb.CreateQuery("SELDREQQR11", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELREQNOQR = rdb.CreateQuery("SELREQNOQRDEP", str)

'**** ADDED FOR BIOCHEMISTRY, EL, TM INTERFACEING
str = "INSERT INTO DISLIB.DBCINFACE(REQNO,LABNO,PROCODE) VALUES(?,?,?)"
Set INSDBCINFACEQR = rdb.CreateQuery("INSDBCINFACEQRDEPT", str)

str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
Set SELDMULPROCQR = rdb.CreateQuery("SELDMULPROCQRBCdept", str)

str = "SELECT * FROM DISLIB.DBCPROSUB WHERE BCPROCODE=?"
Set SELDBCPROSUBQR = rdb.CreateQuery("SELDBCPROSUBQRBCdept", str)

str = "SELECT * FROM DISLIB.DELPROSUB WHERE PROCODE=?"
Set SELELPROSUBQR = rdb.CreateQuery("SELELPROSUBQRELdept", str)

'justin
 str = "SELECT max(pathno) FROM DISLIB.DREQUISITN WHERE REQNO LIKE '%TM%' AND PATHNO LIKE ?"
 Set SELDREQUISITN_PATH_QR = rdb.CreateQuery("1SELDREQUISITN_PATH_QR1", str)

'**** ADDED FOR BIOCHEMISTRY, EL, TM INTERFACEING
End If
MSFlexGrid1.row = 1
MSFlexGrid1.col = 5
chk_rcvd.Width = MSFlexGrid1.CellWidth
chk_rcvd.Height = MSFlexGrid1.CellHeight
chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

MSFlexGrid1.row = 1
MSFlexGrid1.col = 6
txt_pathno.Width = MSFlexGrid1.CellWidth
txt_pathno.Height = MSFlexGrid1.CellHeight
txt_pathno.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top


NEWFLAG = False
viewflag = False
btnenabl ("INI")
FLDENABL ("DIS")
'Opt_all.Enabled = False
'opt_caseno.Enabled = False
'opt_patient.Enabled = False
'opt_staff.Enabled = False
'Opt_all.Value = False
'opt_caseno.Value = False
'opt_patient.Value = False
'opt_staff.Value = False
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub FLDENABL(str1 As String)
If str1 = "ENA" Then
    cmb_dept.Enabled = True
    cmb_modality.Enabled = True
'    Msk_fromdt.Enabled = True
'    Msk_todt.Enabled = True

End If
If str1 = "DIS" Then
    cmb_dept.Enabled = False
    cmb_modality.Enabled = False
'    Msk_fromdt.Enabled = False
   ' Msk_todt.Enabled = False
End If
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

Private Sub MSFlexGrid1_Click()
On Error GoTo MSG:
R1 = MSFlexGrid1.row
C1 = MSFlexGrid1.col
II = req_chglen(MSFlexGrid1.TextMatrix(R1, 1))
If MSFlexGrid1.col = 1 Then
    On Error GoTo MSG:
    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\vpp.RPT"
    CrystalReport1.SelectionFormula = " {RREQUISITN.REQNNO} ='" & req_chglen(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)) & "'"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
    Exit Sub
End If
If MODIFYFLAG = False Then
If MODALITY = "SP" Or MODALITY = "HL" Or MODALITY = "BM" Or MODALITY = "BF" Then
    If MSFlexGrid1.col = 5 Then
    'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then

            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
            'txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))

             chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
             chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            'txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            If Trim(MSFlexGrid1.Text) = "RECEIVED" Then
            chk_rcvd.Value = 1
            chk_rcvd.SetFocus
            Else
            chk_rcvd.Value = 0
            End If

     End If
Else
    If dept <> "NZZ" Then
        If MSFlexGrid1.col = 6 Then
        'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then

            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
            txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))

             'Chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            'If Trim(MSFlexGrid1.Text) = "RECEIVED" Then
            'txt_pathno.SetFocus
'            chk_rcvd.Value = 1
'            chk_rcvd.SetFocus
'            Else
'            chk_rcvd.Value = 0
'            End If
        End If
    End If
    If MODALITY = "BP" Or MODALITY = "SS" Or MODALITY = "HT" Or MODALITY = "ALL" Or MODALITY = "TM" Then
        If MSFlexGrid1.col = 5 Then
        'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then

            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
            'txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))

             chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
             chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            'txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            If Trim(MSFlexGrid1.Text) = "RECEIVED" Then
                chk_rcvd.Value = 1
                chk_rcvd.SetFocus
            Else
                chk_rcvd.Value = 0
            End If
        End If
     End If
     Exit Sub
 End If
Else
    If MODALITY = "SP" Or MODALITY = "HL" Or MODALITY = "BM" Or MODALITY = "BF" Then
        If MSFlexGrid1.col = 5 Then
    'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then

            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
            'txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))

            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            'txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            If Trim(MSFlexGrid1.Text) = "RECEIVED" Then
            chk_rcvd.Value = 1
            Else
            chk_rcvd.Value = 0
            End If
     End If
   Else
  If MSFlexGrid1.col = 6 Then
    'If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then

            'MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
            txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))

            'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            'txt_pathno.SetFocus
'            If Trim(MSFlexGrid1.Text) = "RECEIVED" Then
'            chk_rcvd.Value = 1
'            Else
'            chk_rcvd.Value = 0
'            End If
     End If
 End If
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub MSFlexGrid1_GotFocus()
If MSFlexGrid1.row = 0 Then
Exit Sub
End If
'MSFlexGrid1.Row = 1
'MSFlexGrid1.Col = 5
'chk_rcvd.Width = MSFlexGrid1.CellWidth
'chk_rcvd.Height = MSFlexGrid1.CellHeight
'chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
If dept <> "NZZ" Then
If scrlck_flg = True Then
    MSFlexGrid1.row = 1
End If
MSFlexGrid1.col = 6
 If txt_pathno.Visible <> True Then
    txt_pathno.Width = MSFlexGrid1.CellWidth
    txt_pathno.Height = MSFlexGrid1.CellHeight
    txt_pathno.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
 End If
End If
scrlck_flg = False
End Sub

Private Sub msk_caseno_LostFocus()
On Error GoTo MSG:
CASENO = msk_caseno.Text
If viewflag = True Then
    If Trim(dept) = "GZZ" Then
        cmd_modify.Enabled = True
    End If
End If
If Opt_patient.Value = True Then
    CASENO = padzero(5, "C", msk_caseno.Text, 3)
    msk_caseno.Text = UCase(CASENO)
    If SELTA1011FLAG = False Then
        SELTA1011_QR(0) = msk_caseno.FormattedText
        Set SELTA1011_RS = SELTA1011_QR.OpenResultset(1, 3)
        SELTA1011FLAG = True
    Else
        SELTA1011_QR(0) = msk_caseno.FormattedText
        SELTA1011_RS.MoreResults
        SELTA1011_RS.Requery
    End If
    If SELTA1011_RS.RowCount > 0 Then
        txt_name = Trim(SELTA1011_RS("tag")) & " " & Trim(SELTA1011_RS("first_name")) & " " & Trim(SELTA1011_RS("mid_name")) & " " & Trim(SELTA1011_RS("last_name"))
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
    If dept = "NZZ" And Trim(MODALITY) = "ALL" Then
        DTM1
        Exit Sub
    ElseIf dept = "NZZ" And Trim(MODALITY) <> "ALL" Then
        DTM1MODWISE
        Exit Sub
    End If
    If NEWFLAG = True Then
        'NEWFLAG = False
        If SELRREQNOFLAG = False Then
            SELRREQNO_QR(0) = Trim(dept)
            If SELRREQNO_QR(0) = "FEC" Then
                SELRREQNO_QR(1) = Trim(MODALITY)
                SELRREQNO_QR(2) = msk_caseno.FormattedText
                SELRREQNO_QR(3) = "C"
                SELRREQNO_QR(4) = ""
                SELRREQNO_QR(5) = ""
                SELRREQNO_QR(6) = ""
                Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                'SELRREQNOFLAG = True
            Else
                SELRREQNO_QR(1) = Trim(MODALITY)
                SELRREQNO_QR(2) = msk_caseno.FormattedText
                SELRREQNO_QR(3) = "C"
                SELRREQNO_QR(4) = "P"
                SELRREQNO_QR(5) = "P"
                SELRREQNO_QR(6) = "P"
                Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                SELRREQNOFLAG = True
            End If
        Else
            SELRREQNO_QR(0) = Trim(dept)
            If SELRREQNO_QR(0) = "FEC" Then
                SELRREQNO_QR(1) = Trim(MODALITY)
                SELRREQNO_QR(2) = msk_caseno.FormattedText
                SELRREQNO_QR(3) = "C"
                SELRREQNO_QR(4) = "P"
                SELRREQNO_QR(5) = "P"
                SELRREQNO_QR(6) = "P"
                Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                'SELRREQNO_RS.MoreResults
                'SELRREQNO_RS.Requery
            Else
                SELRREQNO_QR(1) = Trim(MODALITY)
                SELRREQNO_QR(2) = msk_caseno.FormattedText
                SELRREQNO_QR(3) = "C"
                SELRREQNO_QR(4) = "P"
                SELRREQNO_QR(5) = "P"
                SELRREQNO_QR(6) = "P"
                Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                'SELRREQNO_RS.MoreResults
                'SELRREQNO_RS.Requery
            End If
        End If
        If SELRREQNO_RS.RowCount > 0 Then
            i = 1
            MSFlexGrid1.Rows = 1
            While Not SELRREQNO_RS.EOF
                If SELRREQNO_RS("SER_CODE") <> "8888" Then
                    If SELMA101FLAG = False Then
                        SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                        SELMA101FLAG = True
                    Else
                        SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                        SELMA101_RS.MoreResults
                        SELMA101_RS.Requery
                    End If
                    If Trim(MODALITY) = "SP" Then
                        If SELDREQFLAG = False Then
                            SELDREQQR(0) = (SELRREQNO_RS("reqnno"))
                            Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                            SELDREQFLAG = True
                        Else
                            SELDREQQR(0) = (SELRREQNO_RS("reqnno"))
                            SELDREQRS.MoreResults
                            SELDREQRS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            If Trim(SELRREQNO_RS("reqnstat")) = "C" Then
                                If SELDREQRS.RowCount > 0 Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & SELDREQRS("PATHNO")
                                    'chk_collected.Value = 1
                                    'txt_pathno.Visible = False
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & ""
                                    'chk_collected.Value = 1
                                    'txt_pathno.Visible = False
                                End If
                            Else
                                If SELDREQRS.RowCount > 0 Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & Trim(SELDREQRS("PATHNO"))
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & ""
                                End If
                            End If
                            i = i + 1
                        End If
                    Else
                        If SELMA101_RS.RowCount > 0 Then

                            If Trim(SELRREQNO_RS("reqnstat")) = "C" Then
                                If dept = "NZZ" Then
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & ""
                                            GoTo m
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & "" & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & ""
                                    'chk_collected.Value = 1
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                End If
                            Else
                                If dept = "NZZ" Then
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " "
                                            GoTo m
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " "
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det"))
                                End If
                            End If
                            i = i + 1
                        End If
                    End If
                Else
                    If Trim(SELRREQNO_RS("reqnstat")) = "C" Then
                        If dept = "NZZ" Then
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & ""
                                    GoTo m
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & "" & Trim(SELDTMNO_RS("DTMNO"))
                                    GoTo m
                                End If
                            End If
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & ""
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED"
                            'chk_collected.Value = 1
                        End If
                    Else
                        If dept = "NZZ" Then
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " "
                                    GoTo m
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                    GoTo m
                                End If
                            End If
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " "
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE"
                        End If
                    End If
m:
                    i = i + 1
                End If
                SELRREQNO_RS.MoveNext
            Wend

            If Trim(MODALITY) = "SP" Then
                chk_rcvd.Visible = True
                txt_pathno.Visible = False
            Else
                chk_rcvd.Visible = False
                txt_pathno.Visible = True
                txt_pathno.SetFocus
            End If

            If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or MODALITY = "HL" Or MODALITY = "BM" Or MODALITY = "BF" Then 'Or MODALITY = "TM"
                MSFlexGrid1.row = 1
                MSFlexGrid1.col = 5
                chk_rcvd.Width = MSFlexGrid1.CellWidth
                chk_rcvd.Height = MSFlexGrid1.CellHeight
                chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                chk_rcvd.Value = 0
                chk_rcvd.Visible = True
                txt_pathno.Visible = False
            End If
            btnenabl ("NEW")
            'chk_collected.Visible = True
        Else
            MsgBox "No pending Requisition No."
            'chk_rcvd.Visible = False
            txt_pathno.Visible = False
            'btnenabl ("INI")
            FLDENABL ("DIS")
            'cmd_new.SetFocus
            Exit Sub
        End If
        If Trim(MODALITY) = "SP" Then
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 5
            chk_rcvd.Width = MSFlexGrid1.CellWidth
            chk_rcvd.Height = MSFlexGrid1.CellHeight
            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_rcvd.Value = 0
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
        End If
    End If
    If viewflag = True Then
    viewflag = False
    If SELRREQNOFLAG = False Then
                    SELRREQNO_QR(0) = Trim(dept)
                    SELRREQNO_QR(1) = Trim(MODALITY)
                    SELRREQNO_QR(2) = msk_caseno.FormattedText
                    SELRREQNO_QR(3) = "E"
                    SELRREQNO_QR(4) = "E"
                    SELRREQNO_QR(5) = "T"
                    SELRREQNO_QR(6) = "F"
                    Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                    SELRREQNOFLAG = True
                Else
                    SELRREQNO_QR(0) = Trim(dept)
                    SELRREQNO_QR(1) = Trim(MODALITY)
                    SELRREQNO_QR(2) = msk_caseno.FormattedText
                    SELRREQNO_QR(3) = "E"
                    SELRREQNO_QR(4) = "E"
                    SELRREQNO_QR(5) = "T"
                    SELRREQNO_QR(6) = "F"
                    Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                    SELRREQNO_RS.MoreResults
                    SELRREQNO_RS.Requery
                End If
                If SELRREQNO_RS.RowCount > 0 Then
                    i = 1
                    MSFlexGrid1.Rows = 1
                    While Not SELRREQNO_RS.EOF
                        If Trim(SELRREQNO_RS("SER_CODE")) <> "8888" Then
                                If SELMA101FLAG = False Then
                                    SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                                    SELMA101FLAG = True
                                Else
                                    SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                                    SELMA101_RS.MoreResults
                                    SELMA101_RS.Requery
                                End If
                                If SELMA101_RS.RowCount > 0 Then
                                    If SELDREQFLAG = False Then
                                        SELDREQQR(0) = Trim(SELRREQNO_RS("reqnno"))
                                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                        SELDREQFLAG = True
                                    Else
                                        SELDREQQR(0) = Trim(SELRREQNO_RS("reqnno"))
                                        SELDREQRS.MoreResults
                                        SELDREQRS.Requery
                                    End If

                                        If Trim(SELRREQNO_RS("REQNSTAT")) = "E" Then
                                            If SELDREQRS.RowCount > 0 Then
                                                 If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                            i = i + 1
                                                            GoTo n
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            i = i + 1
                                                            GoTo n
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                    i = i + 1
                                                 Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                                                    i = i + 1
                                                 End If
                                            Else
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            i = i + 1
                                                            GoTo n
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            i = i + 1
                                                            GoTo n
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                    i = i + 1
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                    i = i + 1
                                                End If
                                            End If
                                        Else
                                            If Trim(SELRREQNO_RS("REQNSTAT")) = "T" Or Trim(SELRREQNO_RS("REQNSTAT")) = "T" Then
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                            i = i + 1
                                                            GoTo n
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            i = i + 1
                                                            GoTo n
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                    i = i + 1
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                                                    i = i + 1
                                                End If
                                            End If
                                        End If

                                   ' i = i + 1

                                End If
                        Else
                                    If SELDREQFLAG = False Then
                                        SELDREQQR(0) = (SELRREQNO_RS("reqnno"))
                                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                        SELDREQFLAG = True
                                    Else
                                        SELDREQQR(0) = (SELRREQNO_RS("reqnno"))
                                        SELDREQRS.MoreResults
                                        SELDREQRS.Requery
                                    End If

                                    If Trim(SELRREQNO_RS("REQNSTAT")) = "E" Then
                                            If SELDREQRS.RowCount > 0 Then
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                            'i = i + 1
                                                            Exit Sub
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))

                                                            Exit Sub
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                                                End If
                                            Else
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) ' & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            GoTo n
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            GoTo n
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) ' & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                End If
                                            End If
                                        Else
                                            If Trim(SELRREQNO_RS("REQNSTAT")) = "T" Or Trim(SELRREQNO_RS("REQNSTAT")) = "T" Then
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                            GoTo n
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                            GoTo n
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                                                End If
                                            End If
                                        End If
n:
                                     i = i + 1
                        End If
                        SELRREQNO_RS.MoveNext
                    Wend
                    If Trim(MODALITY) = "SP" Then
                        chk_rcvd.Visible = True
                        txt_pathno.Visible = False
                    Else
                        chk_rcvd.Visible = False
                        txt_pathno.Visible = True
                        txt_pathno.SetFocus
                    End If
                    If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Then
                        MSFlexGrid1.row = 1
                        MSFlexGrid1.col = 5
                        chk_rcvd.Width = MSFlexGrid1.CellWidth
                        chk_rcvd.Height = MSFlexGrid1.CellHeight
                        chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                        chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                        chk_rcvd.Value = 0
                        chk_rcvd.Visible = True
                        txt_pathno.Visible = False
                    End If
                           'chk_rcvd.Visible = False
                   txt_pathno.Visible = False
                   'chk_collected.Visible = False
                End If
        End If

End If
If opt_staff.Value = True Then
    STAFF
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub











Private Sub msk_todt_LostFocus()
Dim i As Integer
On Error GoTo MSG:
If OPT_ALL.Value = True Then
    If dept = "NZZ" And Trim(MODALITY) = "ALL" Then
        DTM
        Exit Sub
    ElseIf dept = "NZZ" And Trim(MODALITY) <> "ALL" Then
        DTMMODWISE
        Exit Sub
    End If
    If NEWFLAG = True Then
                    'NEWFLAG = False
                If SELRREQUISITNFLAG = False Then
                    SELRREQUISITN_QR(0) = Trim(dept)
                    If SELRREQUISITN_QR(0) = "FEC" Then
                        SELRREQUISITN_QR(1) = Trim(MODALITY)
                        SELRREQUISITN_QR(2) = msk_fromdt.FormattedText
                        SELRREQUISITN_QR(3) = msk_todt.FormattedText
                        SELRREQUISITN_QR(4) = "C"
                        SELRREQUISITN_QR(5) = ""
                        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                        SELRREQUISITNFLAG = True
                    Else
                        SELRREQUISITN_QR(1) = Trim(MODALITY)
                        SELRREQUISITN_QR(2) = msk_fromdt.FormattedText
                        SELRREQUISITN_QR(3) = msk_todt.FormattedText
                        SELRREQUISITN_QR(4) = "C"
                        SELRREQUISITN_QR(5) = "P"
                        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                        SELRREQUISITNFLAG = True
                    End If
                Else
                        SELRREQUISITN_QR(0) = Trim(dept)
                        If SELRREQUISITN_QR(0) = "FEC" Then
                            SELRREQUISITN_QR(1) = Trim(MODALITY)
                            SELRREQUISITN_QR(2) = msk_fromdt.FormattedText
                            SELRREQUISITN_QR(3) = msk_todt.FormattedText
                            SELRREQUISITN_QR(4) = "C"
                            SELRREQUISITN_QR(5) = ""
                            SELRREQUISITN_RS.MoreResults
                            SELRREQUISITN_RS.Requery
                        Else
                            SELRREQUISITN_QR(1) = Trim(MODALITY)
                            SELRREQUISITN_QR(2) = msk_fromdt.FormattedText
                            SELRREQUISITN_QR(3) = msk_todt.FormattedText
                            SELRREQUISITN_QR(4) = "C"
                            SELRREQUISITN_QR(5) = "P"
                            SELRREQUISITN_RS.MoreResults
                            SELRREQUISITN_RS.Requery
                        End If
                End If

                If SELRREQUISITN_RS.RowCount = 0 Then
                    Exit Sub
                End If
                If SELRREQUISITN_RS.RowCount > 0 Then
                    i = 1
                    MSFlexGrid1.Rows = 1
                    While Not SELRREQUISITN_RS.EOF
                        If SELDTMNO_FLAG = False Then
                            SELDTMNO_QR(0) = SELRREQUISITN_RS("CASENO") ' msk_caseno.FormattedText
                            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                            SELDTMNO_FLAG = True
                        Else
                            SELDTMNO_QR(0) = SELRREQUISITN_RS("CASENO") 'msk_caseno.FormattedText
                            SELDTMNO_RS.MoreResults
                            SELDTMNO_RS.Requery
                        End If
                        If SELRREQUISITN_RS("SER_CODE") <> "8888" Then
                                If SELMA101FLAG = False Then
                                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS("ser_code"))
                                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                                    SELMA101FLAG = True
                                Else
                                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS("ser_code"))
                                    SELMA101_RS.MoreResults
                                    SELMA101_RS.Requery
                                End If
                                If Trim(MODALITY) = "SP" Then
                                        If SELDREQFLAG = False Then
                                                SELDREQQR(0) = (SELRREQUISITN_RS("reqnno"))
                                                Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                                SELDREQFLAG = True
                                        Else
                                                SELDREQQR(0) = (SELRREQUISITN_RS("reqnno"))
                                                SELDREQRS.MoreResults
                                                SELDREQRS.Requery
                                        End If
                                        If SELMA101_RS.RowCount > 0 Then
                                            If Trim(SELRREQUISITN_RS("reqnstat")) = "C" Then
                                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & SELDREQRS("PATHNO")
                                                'chk_collected.Value = 1
                                            Else
                                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "" & vbTab & " " & vbTab & SELDREQRS("PATHNO")
                                            End If
                                            i = i + 1
                                        End If
                                    Else
                                        If SELMA101_RS.RowCount > 0 Then
                                            If Trim(SELRREQUISITN_RS("reqnstat")) = "C" Then
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                                            GoTo m
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                                            GoTo m
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                                End If
                                            'chk_collected.Value = 1
                                            Else
                                                If dept = "NZZ" Then
                                                    If Not SELDTMNO_RS.EOF Then
                                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) '& vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                                            GoTo m
                                                        Else
                                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                                            GoTo m
                                                        End If
                                                    End If
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) '& vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det"))
                                                End If
                                            End If
m:
                                            i = i + 1
                                        End If
                                    End If
                                Else
                                    If Trim(SELRREQUISITN_RS("reqnstat")) = "C" Then
                                        If dept = "NZZ" Then
                                            If Not SELDTMNO_RS.EOF Then
                                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " '& vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                                    GoTo n
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                                    GoTo n
                                                End If
                                            End If
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " '& vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                        'chk_collected.Value = 1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED"
                                        End If
                                    Else
                                        If dept = "NZZ" Then
                                            If Not SELDTMNO_RS.EOF Then
                                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" '& vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                                                    GoTo n
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                                                    GoTo n
                                                End If
                                            End If
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" '& vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE"
                                        End If
                                    End If
n:
                                    i = i + 1
                        End If
                        SELRREQUISITN_RS.MoveNext
                    Wend
                    If Trim(MODALITY) = "SP" Or Trim(MODALITY) = "HL" Or Trim(MODALITY) = "BM" Or Trim(MODALITY) = "BF" Then
                        chk_rcvd.Visible = True
                        txt_pathno.Visible = False
                        chk_rcvd.SetFocus
                    Else
                        chk_rcvd.Visible = False
                        txt_pathno.Visible = True
                    End If
                    'chk_collected.Visible = True
                    If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or Trim(MODALITY) = "HL" Or Trim(MODALITY) = "BM" Or Trim(MODALITY) = "BF" Then
                        MSFlexGrid1.row = 1
                        MSFlexGrid1.col = 5
                        chk_rcvd.Width = MSFlexGrid1.CellWidth
                        chk_rcvd.Height = MSFlexGrid1.CellHeight
                        chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                        chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

                        chk_rcvd.Visible = True
                        txt_pathno.Visible = False
                    End If
                    btnenabl ("NEW")
                Else
                    MsgBox "No pending Requisition No."
                    'chk_rcvd.Visible = False
                    txt_pathno.Visible = False
                    'btnenabl ("INI")
                    FLDENABL ("DIS")
                    'cmd_new.SetFocus
                    Exit Sub
                End If
        End If
        If viewflag = True Then
                viewflag = False
                If SELRREQUISITNFLAG = False Then
                    SELRREQUISITN_QR(0) = Trim(dept)
                    SELRREQUISITN_QR(1) = Trim(MODALITY)
                    SELRREQUISITN_QR(2) = msk_fromdt.FormattedText
                    SELRREQUISITN_QR(3) = msk_todt.FormattedText
                    SELRREQUISITN_QR(4) = "C"
                    SELRREQUISITN_QR(5) = "E"
                    Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
                    SELRREQUISITNFLAG = True
                Else
                    SELRREQUISITN_QR(0) = Trim(dept)
                    SELRREQUISITN_QR(1) = Trim(MODALITY)
                    SELRREQUISITN_QR(2) = msk_fromdt.FormattedText
                    SELRREQUISITN_QR(3) = msk_todt.FormattedText
                    SELRREQUISITN_QR(4) = "C"
                    SELRREQUISITN_QR(5) = "E"
                    SELRREQUISITN_RS.MoreResults
                    SELRREQUISITN_RS.Requery
                End If
                If SELRREQUISITN_RS.RowCount > 0 Then
                    i = 1
                    MSFlexGrid1.Rows = 1
                    While Not SELRREQUISITN_RS.EOF
                        If SELDTMNO_FLAG = False Then
                            SELDTMNO_QR(0) = SELRREQUISITN_RS("CASENO") ' msk_caseno.FormattedText
                            Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                            SELDTMNO_FLAG = True
                        Else
                            SELDTMNO_QR(0) = SELRREQUISITN_RS("CASENO") 'msk_caseno.FormattedText
                            SELDTMNO_RS.MoreResults
                            SELDTMNO_RS.Requery
                        End If
                        If Trim(SELRREQUISITN_RS("SER_CODE")) <> "8888" Then
                                If SELMA101FLAG = False Then
                                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS("ser_code"))
                                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                                    SELMA101FLAG = True
                                Else
                                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS("ser_code"))
                                    SELMA101_RS.MoreResults
                                    SELMA101_RS.Requery
                                End If
                                If SELDREQFLAG = False Then
                                        SELDREQQR(0) = (SELRREQUISITN_RS("reqnno"))
                                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                        SELDREQFLAG = True
                                Else
                                        SELDREQQR(0) = (SELRREQUISITN_RS("reqnno"))
                                        SELDREQRS.MoreResults
                                        SELDREQRS.Requery
                                End If
                                If SELMA101_RS.RowCount > 0 Then
                                    If Trim(SELRREQUISITN_RS("REQNSTAT")) = "E" Then
                                        If SELDREQRS.RowCount > 0 Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn"))
                                        End If
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"

                                    End If
                                    i = i + 1
                                End If
                        Else
                                If SELDREQFLAG = False Then
                                        SELDREQQR(0) = (SELRREQUISITN_RS("reqnno"))
                                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                        SELDREQFLAG = True
                                Else
                                        SELDREQQR(0) = (SELRREQUISITN_RS("reqnno"))
                                        SELDREQRS.MoreResults
                                        SELDREQRS.Requery
                                End If
                                    If Trim(SELRREQUISITN_RS("REQNSTAT")) = "E" Then
                                        If SELDREQRS.RowCount > 0 Then
                                            If dept = "NZZ" Then
                                                If Not SELDTMNO_RS.EOF Then
                                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                        GoTo G
                                                    Else
                                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                        GoTo G
                                                    End If
                                                End If
                                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            Else
                                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                                            End If
                                        Else
                                            If dept = "NZZ" Then
                                                If Not SELDTMNO_RS.EOF Then
                                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                        GoTo G
                                                    Else
                                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                        GoTo G
                                                    End If
                                                End If
                                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            Else
                                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS("examinatn"))
                                            End If
                                        End If
                                    Else
                                        If dept = "NZZ" Then
                                            If Not SELDTMNO_RS.EOF Then
                                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" '& vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                    GoTo G
                                                Else
                                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                                    GoTo G
                                                End If
                                            End If
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" '& vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS("reqnno")) & vbTab & SELRREQUISITN_RS("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED"
                                        End If
                                    End If
G:
                                    i = i + 1
                        End If
                        SELRREQUISITN_RS.MoveNext
                    Wend
                   'chk_rcvd.Visible = False
                   txt_pathno.Visible = False
                End If
        End If
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Private Sub Opt_all_Click()
If OPT_ALL.Value = True Then
    msk_fromdt.Text = DATE11
    msk_todt.Text = DATE11
    OPT_CASENO.Enabled = False
    Opt_patient.Enabled = False
    opt_staff.Enabled = False
    msk_caseno.Enabled = False
    lbl_name.Visible = False
    txt_name.Visible = False
    'txt_pathno.Visible = True
    msk_fromdt.SetFocus
End If

End Sub

Private Sub OPT_CASENO_Click()
If OPT_CASENO.Value = True Then
    OPT_ALL.Value = False
    OPT_ALL.Enabled = False
    msk_fromdt.Enabled = False
    msk_todt.Enabled = False
    'Opt_patient.Value = True
    msk_caseno.Enabled = True
    lbl_name.Visible = True
    txt_name.Visible = True
    msk_caseno.Enabled = True
    msk_caseno.SetFocus
End If
End Sub

Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
    OPT_CASENO.Caption = "Case No."
    msk_caseno.Mask = "??>/#####"
End If
End Sub

Private Sub opt_staff_Click()
If opt_staff.Value = True Then
    OPT_CASENO.Caption = "CC No."
    msk_caseno.Mask = "######"
End If
End Sub

Public Sub STAFF()
        If SELEMPBASINFO_FLAG = False Then
             SELEMPBASINFO_QR(0) = msk_caseno.Text
            Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
             SELEMPBASINFO_FLAG = True
        Else
             SELEMPBASINFO_QR(0) = msk_caseno.Text
             SELEMPBASINFO_RS.MoreResults
             SELEMPBASINFO_RS.Requery
        End If
        If SELEMPBASINFO_RS.RowCount > 0 Then
            txt_name = Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME"))
        End If
        If dept = "NZZ" And Trim(MODALITY) = "ALL" Then
            DTM1
            Exit Sub
        ElseIf dept = "NZZ" And Trim(MODALITY) <> "ALL" Then
            DTM1MODWISE
            Exit Sub
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
'        CASENO = padzero(5, "C", msk_caseno.Text, 3)
'        msk_caseno.Text = UCase(CASENO)
'                If SELTA1011FLAG = False Then
'                    SELTA1011_QR(0) = msk_caseno.FormattedText
'                    Set SELTA1011_RS = SELTA1011_QR.OpenResultset(1, 3)
'                    SELTA1011FLAG = True
'                Else
'                    SELTA1011_QR(0) = msk_caseno.FormattedText
'                    SELTA1011_RS.MoreResults
'                    SELTA1011_RS.Requery
'                End If
'                If SELTA1011_RS.RowCount > 0 Then
'                    txt_name = Trim(SELTA1011_RS("tag")) & " " & Trim(SELTA1011_RS("first_name")) & " " & Trim(SELTA1011_RS("mid_name")) & " " & Trim(SELTA1011_RS("last_name"))
'                End If


                If NEWFLAG = True Then
                'NEWFLAG = False
                If SELRREQNOFLAG = False Then
                    SELRREQNO_QR(0) = Trim(dept)
                    If SELRREQNO_QR(0) = "FEC" Then
                        SELRREQNO_QR(1) = Trim(MODALITY)
                        SELRREQNO_QR(2) = Trim(msk_caseno.FormattedText)
                        SELRREQNO_QR(3) = "C"
                        SELRREQNO_QR(4) = ""
                        SELRREQNO_QR(5) = ""
                        SELRREQNO_QR(6) = ""
                        Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                        SELRREQNOFLAG = True
                    Else
                        SELRREQNO_QR(1) = Trim(MODALITY)
                        SELRREQNO_QR(2) = Trim(msk_caseno.FormattedText)
                        SELRREQNO_QR(3) = "C"
                        SELRREQNO_QR(4) = "P"
                        SELRREQNO_QR(5) = "P"
                        SELRREQNO_QR(6) = "P"
                        Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                        SELRREQNOFLAG = True
                    End If
                Else
                    SELRREQNO_QR(0) = Trim(dept)
                    If SELRREQNO_QR(0) = "FEC" Then
                        SELRREQNO_QR(1) = Trim(MODALITY)
                        SELRREQNO_QR(2) = msk_caseno.FormattedText
                        SELRREQNO_QR(3) = "C"
                        SELRREQNO_QR(4) = ""
                        SELRREQNO_QR(5) = ""
                        SELRREQNO_QR(6) = ""
                        Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                        SELRREQNO_RS.MoreResults
                        SELRREQNO_RS.Requery
                    Else
                        SELRREQNO_QR(1) = Trim(MODALITY)
                        SELRREQNO_QR(2) = msk_caseno.FormattedText
                        SELRREQNO_QR(3) = "C"
                        SELRREQNO_QR(4) = "P"
                        SELRREQNO_QR(5) = "P"
                        SELRREQNO_QR(6) = "P"
                        Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                        SELRREQNO_RS.MoreResults
                        SELRREQNO_RS.Requery
                    End If
                End If
                If SELDTMNO_FLAG = False Then
                    SELDTMNO_QR(0) = Trim(msk_caseno.FormattedText)
                    Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                    SELDTMNO_FLAG = True
                Else
                    SELDTMNO_QR(0) = Trim(msk_caseno.FormattedText)
                    SELDTMNO_RS.MoreResults
                    SELDTMNO_RS.Requery
                End If
                If SELRREQNO_RS.RowCount > 0 Then
                    i = 1
                    MSFlexGrid1.Rows = 1
                    While Not SELRREQNO_RS.EOF
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            If Trim(SELRREQNO_RS("reqnstat")) = "C" Then
                                If dept = "NZZ" Then
                                    If SELDTMNO_RS.EOF Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & SELDTMNO_RS("dtmno")
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                'chk_collected.Value = 1
                            Else
                                If dept = "NZZ" Then
                                    If SELDTMNO_RS.EOF Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det"))
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & SELDTMNO_RS("dtmno")
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det"))
                            End If
                            i = i + 1
                        End If
                        SELRREQNO_RS.MoveNext
                    Wend
                    'NEW LINE ADDED
                    If Trim(MODALITY) = "SP" Then
                        chk_rcvd.Visible = True
                        txt_pathno.Visible = False
                    Else
                        chk_rcvd.Visible = False
                        txt_pathno.Visible = True
                    End If
                    'NEW LINE ADDED

                    'If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Then
                    ' replaced from above code as  all mod were not there sundeep on 16-02-2004

                    If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or MODALITY = "TM" Or MODALITY = "HL" Or MODALITY = "BM" Or MODALITY = "BF" Then
                        MSFlexGrid1.row = 1
                        MSFlexGrid1.col = 5
                        chk_rcvd.Width = MSFlexGrid1.CellWidth
                        chk_rcvd.Height = MSFlexGrid1.CellHeight
                        chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                        chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

                        chk_rcvd.Visible = True
                        txt_pathno.Visible = False
''                   ' Else    'COMMENTED
                        'chk_rcvd.Visible = True
''                   '     txt_pathno.Visible = True
                    End If
                    'chk_collected.Visible = True
                Else
                    MsgBox "No pending Requisition No."
                    'chk_rcvd.Visible = False
                    txt_pathno.Visible = False
                    'btnenabl ("INI")
                    FLDENABL ("DIS")
                    'cmd_new.SetFocus
                    Exit Sub
                End If
        End If
        If viewflag = True Then
                viewflag = False
                If SELRREQNOFLAG = False Then
                    SELRREQNO_QR(0) = Trim(dept)
                    SELRREQNO_QR(1) = Trim(MODALITY)
                    SELRREQNO_QR(2) = msk_caseno.FormattedText
                    SELRREQNO_QR(3) = "C"
                    SELRREQNO_QR(4) = "E"
                    SELRREQNO_QR(5) = "T"
                    SELRREQNO_QR(6) = "F"
                    Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                    SELRREQNOFLAG = True
                Else
                    SELRREQNO_QR(0) = Trim(dept)
                    SELRREQNO_QR(1) = Trim(MODALITY)
                    SELRREQNO_QR(2) = msk_caseno.FormattedText
                    SELRREQNO_QR(3) = "C"
                    SELRREQNO_QR(4) = "E"
                    SELRREQNO_QR(5) = "T"
                    SELRREQNO_QR(6) = "F"
                    Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                    SELRREQNO_RS.MoreResults
                    SELRREQNO_RS.Requery
                End If
                If SELRREQNO_RS.RowCount > 0 Then
                    i = 1
                    MSFlexGrid1.Rows = 1
                    While Not SELRREQNO_RS.EOF
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = Trim(SELRREQNO_RS("ser_code"))
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELDREQFLAG = False Then
                            SELDREQQR(0) = (SELRREQNO_RS("reqnno"))
                            Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                            SELDREQFLAG = True
                        Else
                            SELDREQQR(0) = (SELRREQNO_RS("reqnno"))
                            SELDREQRS.MoreResults
                            SELDREQRS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            If Trim(SELRREQNO_RS("REQNSTAT")) = "E" Then
                                If dept = "NZZ" Then
                                    If SELDTMNO_RS.EOF Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn")) & vbTab & SELDTMNO_RS("dtmno")
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS("examinatn"))
                            Else
                                If dept = "NZZ" Then
                                    If SELDTMNO_RS.EOF Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & SELDTMNO_RS("dtmno")
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                            End If
                            i = i + 1
                        End If
                        SELRREQNO_RS.MoveNext
                    Wend
                   'chk_rcvd.Visible = False
                   txt_pathno.Visible = False
                   'chk_collected.Visible = False
                End If
        End If

End Sub
'ADDED BY AMIT FOR INTERFACES

Private Sub PRINTBARCODE_Click()
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
    'Dim QTY1 As Integer
    Nos = 0
    j = 1

     f1 = FreeFile()
     Open "c:\TATATM.TXT" For Output As #f1

     'FOR NEW FORMAT OF BARCODE PRINTING PLEASE USE
     'FOLLOWING OPEN COMMAND (JUST REMOVE THE COMMENT) BY MILIND
     'Open "c:\tata11.TXT" For Output As #f1
     If BARCODE_NEW = True Then
     SELTIME_QR(0) = REQNO
     End If
     If BARCODE_VIEW = True Then
        SELTIME_QR(0) = II
    End If
     SELTIME_QR.Execute
     Set SELTIME_RS = SELTIME_QR.OpenResultset(1, 3)
     If SELTIME_RS.RowCount > 0 Then
        REC_DATE = Format(SELTIME_RS!SAMRECDATE, "DD/MM/YYYY")
        REC_TIME = Format(SELTIME_RS!SAMRECTIME, "hh:mm:ss")
     Else
        MsgBox "YOU CANNOT PRINT BARCODE LABEL FOR THIS REQUISITN B'COZE NO REQNO EXIST IN RTIMELOG FILE"
        Close #f1
        Exit Sub
    End If
            If BARCODE_NEW = True Then
                SELLABNO_QR(0) = REQNO ' req_chglen(cmb_reqno.Text)
                SELLABNO_QR.Execute
                QTY = 1
                Nos = 0
                Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
                Do Until SELLABNO_RS.EOF
                    BAR_CD = Trim(SELLABNO_RS!LABNO)
                    SELPATH_QR(0) = REQNO
                    SELPATH_QR.Execute
                    Set SELPATH_RS = SELPATH_QR.OpenResultset(1, 3)
                    LABNO = Mid(SELPATH_RS("PATHNO"), 7)

                    Dim count As Integer
                    SELPROCODE_QR(0) = REQNO
                    SELPROCODE_QR.Execute
                    Set SELPROCODE_RS = SELPROCODE_QR.OpenResultset(1, 3)
                    count = SELPROCODE_RS.RowCount


                    'ITEM_DESC = CASENO
                    CASENO = msk_caseno.FormattedText
                     'REC_DATE = txt_date
                     'REC_TIME = txt_time
                     PATNAME = Trim(txt_name.Text)
                    COLL_CENTER = ""
                    QTY1 = "000" & QTY
                    EXP_DT = "EXP" '& Trim(SELLABNO_RS!TDATE) & Trim(SELLABNO_RS!TTIME)

                    Call Printlbl(BAR_CD)
                    'BY MILIND FOR NEW FORMAT OF BARCODE PRINTING PLEASE USE
                    'FOLLOWING CALL COMMAND (JUST REMOVE THE COMMENT, ALL THE CODING IS DONE)
                    'Call PRINTLBLNEW(BAR_CD)
                    Nos = Nos + QTY
                    SELLABNO_RS.MoveNext
                Loop
                Close #f1
                Shell "c:\TATATM.BAT", 0

                'FOR NEW FORMAT OF BARCODE PRINTING PLEASE USE
                'FOLLOWING SHELL COMMAND (REMOVE THE COMMENT) BY MILIND
                'Shell "c:\tata11.bat", 0

                MsgBox "Total " & Nos & " Labels Have Been Printed", vbInformation
            ElseIf BARCODE_VIEW = True Then
                    If II = "" Then
                        MsgBox "PLEASE CLICK ON REQUISITN NO."
                        Close #f1
                        Exit Sub
                    Else
                    SELLABNO_QR(0) = II ' req_chglen(cmb_reqno.Text)
                    SELLABNO_QR.Execute
                    QTY = 1
                    Nos = 0
                    Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
                    Do Until SELLABNO_RS.EOF
                    BAR_CD = Trim(SELLABNO_RS!LABNO)

                    SELPATH_QR(0) = II
                    SELPATH_QR.Execute
                    Set SELPATH_RS = SELPATH_QR.OpenResultset(1, 3)
                    LABNO = Mid(SELPATH_RS("PATHNO"), 7)



                    'ITEM_DESC = CASENO
                    CASENO = msk_caseno.FormattedText
                    'REC_DATE = txt_date
                    'REC_TIME = txt_time
                    PATNAME = Trim(txt_name.Text)
                    COLL_CENTER = ""
                    QTY1 = "000" & QTY
                    EXP_DT = "EXP" '& Trim(SELLABNO_RS!TDATE) & Trim(SELLABNO_RS!TTIME)
                    'Call PRINTLBLNEW(BAR_CD)
                    Call Printlbl(BAR_CD)
                    Nos = Nos + QTY
                    SELLABNO_RS.MoveNext
                    Loop
                    End If
                    Close #f1
                    Shell "c:\TATATM.bat", 0
                    'Shell "c:\tata11.bat", 0
                    MsgBox "Total " & Nos & " Labels Have Been Printed", vbInformation
            End If
     Close #f1

End Sub
'ADDED BY AMIT FOR INTERFACES
Public Sub Printlbl(BAR_CD1 As String) ', caseno As String, patname As String, coll_centre As String, REC_DATE As String, rec_time)
''BAR_CD1 = BAR_CD
BAR_CD1 = Trim(BAR_CD1)
bar1 = Mid(BAR_CD1, 1, 2)
bar2 = Mid(BAR_CD1, 3, 8)
bar3 = Mid(BAR_CD1, 11, 1)
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
Print #f1, "{RB00;" + BAR_CD1 + "|}"
Print #f1, "{RC001;" + BAR_CD1 + "|}"
Print #f1, "{RC002;FEC|}"
Print #f1, "{RC003;" + CASENO + "|}"
Print #f1, "{RC004;" + LABNO + "|}"
Print #f1, "{XS;I,0001,0002C4201|}"


'upto here

' TEC BARCODE PRINTER PAGE END
' <xpml></page></xpml>TEC BARCODE PRINTER JOB END
' <xpml><end/></xpml>


''' NEW CODES





'''' NEW LABEL END


End Sub

Private Sub Timer1_Timer()
datetime
txt_time.Text = time1
End Sub

Private Sub txt_pathno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
    If MSFlexGrid1.row <> 1 Then
    MSFlexGrid1.row = MSFlexGrid1.row - 1
    txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
    'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = ""

    End If
End If
If KeyCode = 40 Then
    If MSFlexGrid1.row <> MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.row = MSFlexGrid1.row + 1
        txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
    End If
End If
End Sub

Private Sub TXT_PATHNO_KeyPress(KeyAscii As Integer)
Dim i As Integer
Dim S As String
Dim ST As String
On Error GoTo MSG:
If KeyAscii = 13 Then
ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
'*****
If NEWFLAG = True Then


    If Trim(txt_pathno.Text) <> "" Then
        If MODALITY = "BC" Or MODALITY = "EL" Or MODALITY = "TM" Then
            If seldequisitnflag = False Then
                If Len(Trim(txt_pathno.Text)) <= 3 And Len(Trim(txt_pathno.Text)) <> 0 Then
                SELDREQUISITNQR(0) = ST & UCase(Trim(txt_pathno.Text))
                End If
                If Len(Trim(txt_pathno.Text)) > 3 Then
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                End If
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                Set SELDREQUISITNRS = SELDREQUISITNQR.OpenResultset(1, 3)
                seldequisitnflag = True
            Else
                If Len(Trim(txt_pathno.Text)) <= 3 Or Len(Trim(txt_pathno.Text)) <> 0 Then
                SELDREQUISITNQR(0) = ST & UCase(Trim(txt_pathno.Text))
                End If
                If Len(Trim(txt_pathno.Text)) > 3 Then
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                End If
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                SELDREQUISITNRS.MoreResults
                SELDREQUISITNRS.Requery
            End If
        Else
            If seldequisitnflag = False Then
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                Set SELDREQUISITNRS = SELDREQUISITNQR.OpenResultset(1, 3)
                seldequisitnflag = True
            Else
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                SELDREQUISITNRS.MoreResults
                SELDREQUISITNRS.Requery
            End If
        End If
        If MODIFYFLAG = False Then
            If SELDREQUISITNRS.RowCount > 0 Then
                MsgBox "The Path No. is  Given already " & SELDREQUISITNRS("caseno")
                'txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
                'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
                'txt_pathno.Visible = True
                'chk_rcvd.Visible = True
                'chk_rcvd.SetFocus
                txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                Exit Sub

            End If
        Else
            MODIFYFLAG = True
        End If
    End If

End If
'*****

If Trim(txt_pathno.Text) <> "" Then
  'trim(MSFlexGrid1.TextMatrix(i, 6) = ""
For i = 1 To MSFlexGrid1.Rows - 1
   If NEWFLAG = True Then
    If Trim(txt_pathno.Text) = Trim(MSFlexGrid1.TextMatrix(i, 6)) Then
         MsgBox "Lab No. Already exists"
         Exit Sub
    End If
   End If
   If MODIFYFLAG = True Then
    '*****
    If MODALITY = "BC" Or MODALITY = "EL" Or MODALITY = "TM" Then
            If seldequisitnflag = False Then
                If Len(Trim(txt_pathno.Text)) <= 3 And Len(Trim(txt_pathno.Text)) <> 0 Then
                SELDREQUISITNQR(0) = ST & UCase(Trim(txt_pathno.Text))
                End If
                If Len(Trim(txt_pathno.Text)) > 3 Then
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                End If
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                Set SELDREQUISITNRS = SELDREQUISITNQR.OpenResultset(1, 3)
                seldequisitnflag = True
            Else
                If Len(Trim(txt_pathno.Text)) <= 3 And Len(Trim(txt_pathno.Text)) <> 0 Then
                SELDREQUISITNQR(0) = ST & UCase(Trim(txt_pathno.Text))
                End If
                If Len(Trim(txt_pathno.Text)) > 3 Then
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                End If
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                SELDREQUISITNRS.MoreResults
                SELDREQUISITNRS.Requery
            End If
        Else
            If seldequisitnflag = False Then
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                Set SELDREQUISITNRS = SELDREQUISITNQR.OpenResultset(1, 3)
                seldequisitnflag = True
            Else
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
                SELDREQUISITNQR(1) = dept & MODALITY & "%"
                SELDREQUISITNRS.MoreResults
                SELDREQUISITNRS.Requery
            End If
        End If

        If SELDREQUISITNRS.RowCount > 0 Then
                MsgBox "The Path No. is  Given already " & SELDREQUISITNRS("caseno")

                txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top

                Exit Sub

            End If


   '****

         If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = MSFlexGrid1.TextMatrix(i, 6) Then
         MSFlexGrid1.TextMatrix(i, 6) = ""
         Else
         S = CStr(MSFlexGrid1.TextMatrix(i, 6))
         End If
        If Trim(txt_pathno.Text) = S Then
         MsgBox "Lab No. Already exists"
         Exit Sub
        Else

    End If
   End If
Next i

MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = "COLLECTED"
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = "RECEIVED"
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = txt_pathno.Text
Else
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4) = ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5) = ""
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = txt_pathno.Text


End If
'If Len(Trim(txt_positive)) = 0 Then
'     cmd_save.SetFocus
'     Exit Sub
'
'End If
'
    i = MSFlexGrid1.Rows - 1
    If i = MSFlexGrid1.row Then
        cmd_save.SetFocus
        Exit Sub
    End If
    If MODIFYFLAG = False Then
        If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
        End If
        If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.row = MSFlexGrid1.row + 1
                txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
                'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
                txt_pathno.Visible = True
                'chk_rcvd.Visible = True
                'chk_rcvd.SetFocus
                txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = ""
        End If
    Else
        If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
        End If
        If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then
                MSFlexGrid1.row = MSFlexGrid1.row + 1
                txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
                'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
                txt_pathno.Visible = True
                'chk_rcvd.Visible = True
                'chk_rcvd.SetFocus
                txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        End If
    End If
    Exit Sub
MSG:
    MsgBox Err.Description
    End If
End Sub

Private Sub txt_pathno_Change()
If MODIFYFLAG = True Then
  MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6) = Trim(txt_pathno.Text)
End If
End Sub

Private Sub txt_pathno_LostFocus()

''If MODIFYFLAG = True Then
''If Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)) <> "" Then
''    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = txt_pathno
''End If
''End If
''Dim i As Integer
'On Error GoTo msg:
''
''MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6) = txt_pathno.Text
'''If Len(Trim(txt_positive)) = 0 Then
'''     cmd_save.SetFocus
'''     Exit Sub
'''
'''End If
'''
''i = MSFlexGrid1.Rows - 1
''If i = MSFlexGrid1.Row Then
''    cmd_save.SetFocus
''    Exit Sub
''End If
''If MODIFYFLAG = False Then
''    If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
''            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
''            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
''    End If
''    If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
''            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
''            txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
''            chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
''            txt_pathno.Visible = True
''            chk_rcvd.Visible = True
''            'chk_rcvd.SetFocus
''            txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''    End If
''Else
''    If MSFlexGrid1.Row = MSFlexGrid1.Rows - 1 Then
''            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
''            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
''    End If
''    If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
''            MSFlexGrid1.Row = MSFlexGrid1.Row + 1
''            txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))
''            'chk_rcvd = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
''            txt_pathno.Visible = True
''            'chk_rcvd.Visible = True
''            'chk_rcvd.SetFocus
''            txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''            'chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''    End If
''End If
''Exit Sub
''msg:
''MsgBox Err.Description
End Sub

Public Sub BCINF()
''ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
''If seldmulprocflag = False Then
''     SELDMULPROCQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''     Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
''     seldmulprocflag = True
''Else
''     SELDMULPROCQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''     seldmulprocrs.MoreResults
''     seldmulprocrs.Requery
''End If
''If seldmulprocrs.RowCount > 0 Then
''
''     While Not seldmulprocrs.EOF
''           If seldmulprocrs("pro_code") = "F849" Or seldmulprocrs("PRO_CODE") = "F854" Or seldmulprocrs("PRO_CODE") = "F860" Or seldmulprocrs("PRO_CODE") = "F865" Or seldmulprocrs("PRO_CODE") = "F875" Or seldmulprocrs("PRO_CODE") = "F880" Or seldmulprocrs("PRO_CODE") = "F802" Then
''
''                    If SELDBCPROSUBFLAG = False Then
''                        SELDBCPROSUBQR(0) = seldmulprocrs("pro_code")
''                        Set SELDBCPROSUBRS = SELDBCPROSUBQR.OpenResultset(1, 3)
''                        SELDBCPROSUBFLAG = True
''                    Else
''                        SELDBCPROSUBQR(0) = seldmulprocrs("pro_code")
''                        SELDBCPROSUBRS.MoreResults
''                        SELDBCPROSUBRS.Requery
''                    End If
''
''                    If SELDBCPROSUBRS.RowCount > 0 Then
''                        While Not SELDBCPROSUBRS.EOF
''                            INSDBCINFACEQR(0) = req_chglen(MSFlexGrid1.TextMatrix(sRow, 1))
''                            If Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))) > 3 Then
''                                INSDBCINFACEQR(1) = MSFlexGrid1.TextMatrix(srow, 6)
''                            End If
''                            If Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) > 0 Then
''                                INSDBCINFACEQR(1) =st & MSFlexGrid1.TextMatrix(srow, 6)
''                            End If
''                            INSDBCINFACEQR(2) = SELDBCPROSUBRS("procodesub")
''                            INSDBCINFACEQR.Execute
''
''                           SELDBCPROSUBRS.MoveNext
''                        Wend
''                 End If
''           Else
''
''                          INSDBCINFACEQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''                           If Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))) > 3 Then
''                                INSDBCINFACEQR(1) = MSFlexGrid1.TextMatrix(srow, 6)
''                            End If
''                            If Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) > 0 Then
''                                INSDBCINFACEQR(1) = st & MSFlexGrid1.TextMatrix(srow, 6)
''                            End If
''                         INSDBCINFACEQR(2) = seldmulprocrs("procodesub")
''                         INSDBCINFACEQR.Execute
''           End If
''        seldmulprocrs.MoveNext
''     Wend
''End If
End Sub

Public Sub ELINF()
''If seldmulprocflag = False Then
''     SELDMULPROCQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''     Set seldmulprocrs = SELDMULPROCQR.OpenResultset(1, 3)
''     seldmulprocflag = True
''Else
''     SELDMULPROCQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''     seldmulprocrs.MoreResults
''     seldmulprocrs.Requery
''End If
''If seldmulprocrs.RowCount > 0 Then
''    While Not seldmulprocrs.EOF
''       If seldmulprocrs("PRO_CODE") = "F860" Or seldmulprocrs("PRO_CODE") = "F875" Or seldmulprocrs("PRO_CODE") = "F901" Or seldmulprocrs("PRO_CODE") = "F903" Or seldmulprocrs("PRO_CODE") = "F908" Then
''            If SELELPROSUBFLAG = False Then
''                SELELPROSUBQR(0) = Trim(seldmulprocrs("PRO_CODE"))
''                Set SELELPROSUBRS = SELELPROSUBQR.OpenResultset(1, 3)
''                SELELPROSUBFLAG = True
''            Else
''                SELELPROSUBQR(0) = Trim(seldmulprocrs("PRO_CODE"))
''                SELELPROSUBRS.MoreResults
''                SELELPROSUBRS.Requery
''            End If
''            If SELELPROSUBRS.RowCount > 0 Then
''                While Not SELELPROSUBRS.EOF
''                        INSDBCINFACEQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''                        'If Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) > 3 Then
''                            INSDBCINFACEQR(1) = MSFlexGrid1.TextMatrix(srow, 6)
''                        'End If
''                        'If Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) > 0 Then
''                         '   INSDBCINFACEQR(1) = ST & MSFlexGrid1.TextMatrix(srow, 6)
''                        'End If
''                        INSDBCINFACEQR(2) = SELELPROSUBRS("procodesub")
''                        INSDBCINFACEQR.Execute
''
''                    SELELPROSUBRS.MoveNext
''                Wend
''            End If
''       Else
''            INSDBCINFACEQR(0) = req_chglen(MSFlexGrid1.TextMatrix(srow, 1))
''            'If Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) > 3 Then
''                INSDBCINFACEQR(1) = MSFlexGrid1.TextMatrix(srow, 6)
''            'End If
''            'If Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(srow, 6))) > 0 Then
''            '    INSDBCINFACEQR(1) = ST & MSFlexGrid1.TextMatrix(srow, 6)
''            'End If
''            INSDBCINFACEQR(2) = seldmulprocrs("PRO_CODE")
''            INSDBCINFACEQR.Execute
''
''       End If
''       seldmulprocrs.MoveNext
''   Wend
''End If
End Sub

Public Sub TMINF()

End Sub
Public Sub DTM()

If NEWFLAG = True Then
    If SELRREQUISITNFLAG1 = False Then
        SELRREQUISITN_QR1(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR1(1) = msk_todt.FormattedText
        SELRREQUISITN_QR1(2) = "C"
        SELRREQUISITN_QR1(3) = "P"
        SELRREQUISITN_QR1(4) = "C"

        Set SELRREQUISITN_RS1 = SELRREQUISITN_QR1.OpenResultset(1, 3)
        SELRREQUISITNFLAG1 = True
    Else
        SELRREQUISITN_QR1(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR1(1) = msk_todt.FormattedText
        SELRREQUISITN_QR1(2) = "C"
        SELRREQUISITN_QR1(3) = "P"
        SELRREQUISITN_QR1(4) = "C"
        SELRREQUISITN_RS1.MoreResults
        SELRREQUISITN_RS1.Requery
    End If
    If SELRREQUISITN_RS1.RowCount = 0 Then
        Exit Sub
    End If
    If SELRREQUISITN_RS1.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQUISITN_RS1.EOF
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = SELRREQUISITN_RS1("CASENO") ' msk_caseno.FormattedText
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = SELRREQUISITN_RS1("CASENO") 'msk_caseno.FormattedText
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELRREQUISITN_RS1("SER_CODE") <> "8888" Then
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS1("ser_code"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS1("ser_code"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If Trim(MODALITY) = "SP" Then
                    If SELDREQFLAG = False Then
                        SELDREQQR(0) = (SELRREQUISITN_RS1("reqnno"))
                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                        SELDREQFLAG = True
                    Else
                        SELDREQQR(0) = (SELRREQUISITN_RS1("reqnno"))
                        SELDREQRS.MoreResults
                        SELDREQRS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQUISITN_RS1("reqnstat")) = "C" Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & SELDREQRS("PATHNO")
                            'chk_collected.Value = 1
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "" & vbTab & " " & vbTab & SELDREQRS("PATHNO")
                        End If
                        i = i + 1
                    End If
                Else
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQUISITN_RS1("reqnstat")) = "C" Then
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                End If
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                            End If
                        Else
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) '& vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                End If
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det"))
                            End If
                        End If
                        i = i + 1
                    End If
                End If
            Else
                If Trim(SELRREQUISITN_RS1("reqnstat")) = "C" Then
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " '& vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED"
                    End If
                Else
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" '& vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE"
                    End If
                End If
                i = i + 1
            End If
            SELRREQUISITN_RS1.MoveNext
        Wend
        If Trim(MODALITY) = "SP" Then
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
            chk_rcvd.SetFocus
        Else
            chk_rcvd.Visible = False
            txt_pathno.Visible = True
        End If
        'chk_collected.Visible = True
        If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or Trim(MODALITY) = "ALL" Then
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 5
            chk_rcvd.Width = MSFlexGrid1.CellWidth
            chk_rcvd.Height = MSFlexGrid1.CellHeight
            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
        End If
    Else
        MsgBox "No pending Requisition No."
        'chk_rcvd.Visible = False
        txt_pathno.Visible = False
        'btnenabl ("INI")
        FLDENABL ("DIS")
        'cmd_new.SetFocus
        Exit Sub
    End If
    MSFlexGrid1.row = 1
    MSFlexGrid1.col = 5
    chk_rcvd.Width = MSFlexGrid1.CellWidth
    chk_rcvd.Height = MSFlexGrid1.CellHeight
    chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    chk_rcvd.Visible = True
    txt_pathno.Visible = False
End If
If viewflag = True Then
    viewflag = False
    If SELRREQUISITNFLAG1 = False Then
        SELRREQUISITN_QR1(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR1(1) = msk_todt.FormattedText
        SELRREQUISITN_QR1(2) = "E"
        SELRREQUISITN_QR1(3) = "T"
        SELRREQUISITN_QR1(4) = "F"
        Set SELRREQUISITN_RS1 = SELRREQUISITN_QR1.OpenResultset(1, 3)
        SELRREQUISITNFLAG1 = True
    Else
        SELRREQUISITN_QR1(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR1(1) = msk_todt.FormattedText
        SELRREQUISITN_QR1(2) = "E"
        SELRREQUISITN_QR1(3) = "T"
        SELRREQUISITN_QR1(4) = "F"
        SELRREQUISITN_RS1.MoreResults
        SELRREQUISITN_RS1.Requery
    End If
    If SELRREQUISITN_RS1.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQUISITN_RS1.EOF
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = SELRREQUISITN_RS1("CASENO") ' msk_caseno.FormattedText
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = SELRREQUISITN_RS1("CASENO") 'msk_caseno.FormattedText
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If Trim(SELRREQUISITN_RS1("SER_CODE")) <> "8888" Then
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS1("ser_code"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS1("ser_code"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If SELDREQFLAG = False Then
                    SELDREQQR(0) = (SELRREQUISITN_RS1("reqnno"))
                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                    SELDREQFLAG = True
                Else
                    SELDREQQR(0) = (SELRREQUISITN_RS1("reqnno"))
                    SELDREQRS.MoreResults
                    SELDREQRS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    If Trim(SELRREQUISITN_RS1("REQNSTAT")) = "E" Then
                        If SELDREQRS.RowCount > 0 Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS1("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS1("examinatn"))
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                    End If
                    i = i + 1
                End If
            Else
                If SELDREQFLAG = False Then
                    SELDREQQR(0) = (SELRREQUISITN_RS1("reqnno"))
                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                    SELDREQFLAG = True
                Else
                    SELDREQQR(0) = (SELRREQUISITN_RS1("reqnno"))
                    SELDREQRS.MoreResults
                    SELDREQRS.Requery
                End If
                If Trim(SELRREQUISITN_RS1("REQNSTAT")) = "E" Then
                    If SELDREQRS.RowCount > 0 Then
                        If Not SELDTMNO_RS.EOF Then
                            If IsNull(SELDTMNO_RS("dtmno")) Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS1("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            End If
                        End If
                    Else
                        If Not SELDTMNO_RS.EOF Then
                            If IsNull(SELDTMNO_RS("dtmno")) Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS1("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            End If
                        End If
                    End If
                Else
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("dtmno")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" '& vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS1("reqnno")) & vbTab & SELRREQUISITN_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                        End If
                    End If
                End If
                i = i + 1
            End If
            SELRREQUISITN_RS1.MoveNext
        Wend
        'chk_rcvd.Visible = False
        txt_pathno.Visible = False
    End If
End If
End Sub
Public Sub DTMMODWISE()

If NEWFLAG = True Then
    If SELRREQUISITN_FLAG11 = False Then
        SELRREQUISITN_QR11(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR11(1) = msk_todt.FormattedText
        SELRREQUISITN_QR11(2) = MODALITY
        SELRREQUISITN_QR11(3) = "C"
        SELRREQUISITN_QR11(4) = "P"
        SELRREQUISITN_QR11(5) = "C"
        Set SELRREQUISITN_RS11 = SELRREQUISITN_QR11.OpenResultset(1, 3)
        SELRREQUISITN_FLAG11 = True
    Else
        SELRREQUISITN_QR11(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR11(1) = msk_todt.FormattedText
        SELRREQUISITN_QR11(2) = MODALITY
        SELRREQUISITN_QR11(3) = "C"
        SELRREQUISITN_QR11(4) = "P"
        SELRREQUISITN_QR11(5) = "C"
        SELRREQUISITN_RS11.MoreResults
        SELRREQUISITN_RS11.Requery
    End If
    If SELRREQUISITN_RS11.RowCount = 0 Then
        Exit Sub
    End If
    If SELRREQUISITN_RS11.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQUISITN_RS11.EOF
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = SELRREQUISITN_RS11("CASENO") ' msk_caseno.FormattedText
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = SELRREQUISITN_RS11("CASENO") 'msk_caseno.FormattedText
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If SELRREQUISITN_RS11("SER_CODE") <> "8888" Then
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS11("ser_code"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS11("ser_code"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If Trim(MODALITY) = "SP" Then
                    If SELDREQFLAG = False Then
                        SELDREQQR(0) = (SELRREQUISITN_RS11("reqnno"))
                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                        SELDREQFLAG = True
                    Else
                        SELDREQQR(0) = (SELRREQUISITN_RS11("reqnno"))
                        SELDREQRS.MoreResults
                        SELDREQRS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQUISITN_RS11("reqnstat")) = "C" Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & SELDREQRS("PATHNO")
                            'chk_collected.Value = 1
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "" & vbTab & " " & vbTab & SELDREQRS("PATHNO")
                        End If
                        i = i + 1
                    End If
                Else
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQUISITN_RS11("reqnstat")) = "C" Then
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                End If
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                            End If
                        Else
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) '& vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                                End If
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det"))
                            End If
                        End If
                        i = i + 1
                    End If
                End If
            Else
                If Trim(SELRREQUISITN_RS11("reqnstat")) = "C" Then
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " '& vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & " " & vbTab & "" & SELDTMNO_RS("DTMNO")
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED"
                    End If
                Else
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" '& vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " " & vbTab & " " & SELDTMNO_RS("DTMNO")
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE"
                    End If
                End If
                i = i + 1
            End If
            SELRREQUISITN_RS11.MoveNext
        Wend
        If Trim(MODALITY) = "SP" Then
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
            chk_rcvd.SetFocus
        Else
            chk_rcvd.Visible = False
            txt_pathno.Visible = True
        End If
        'chk_collected.Visible = True
        If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or Trim(MODALITY) = "ALL" Then
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 5
            chk_rcvd.Width = MSFlexGrid1.CellWidth
            chk_rcvd.Height = MSFlexGrid1.CellHeight
            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
        End If
    Else
        MsgBox "No pending Requisition No."
        'chk_rcvd.Visible = False
        txt_pathno.Visible = False
        'btnenabl ("INI")
        FLDENABL ("DIS")
        'cmd_new.SetFocus
        Exit Sub
    End If
    MSFlexGrid1.row = 1
    MSFlexGrid1.col = 5
    chk_rcvd.Width = MSFlexGrid1.CellWidth
    chk_rcvd.Height = MSFlexGrid1.CellHeight
    chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    chk_rcvd.Visible = True
    txt_pathno.Visible = False
End If
If viewflag = True Then
    viewflag = False
    If SELRREQUISITN_FLAG11 = False Then
        SELRREQUISITN_QR11(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR11(1) = msk_todt.FormattedText
        SELRREQUISITN_QR11(2) = MODALITY
        SELRREQUISITN_QR11(3) = "E"
        SELRREQUISITN_QR11(4) = "T"
        SELRREQUISITN_QR11(5) = "F"
        Set SELRREQUISITN_RS11 = SELRREQUISITN_QR11.OpenResultset(1, 3)
        SELRREQUISITN_FLAG11 = True
    Else
        SELRREQUISITN_QR11(0) = msk_fromdt.FormattedText
        SELRREQUISITN_QR11(1) = msk_todt.FormattedText
        SELRREQUISITN_QR11(2) = MODALITY
        SELRREQUISITN_QR11(3) = "E"
        SELRREQUISITN_QR11(4) = "T"
        SELRREQUISITN_QR11(5) = "F"
        SELRREQUISITN_RS11.MoreResults
        SELRREQUISITN_RS11.Requery
    End If
    If SELRREQUISITN_RS11.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQUISITN_RS11.EOF
            If SELDTMNO_FLAG = False Then
                SELDTMNO_QR(0) = SELRREQUISITN_RS11("CASENO") ' msk_caseno.FormattedText
                Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
                SELDTMNO_FLAG = True
            Else
                SELDTMNO_QR(0) = SELRREQUISITN_RS11("CASENO") 'msk_caseno.FormattedText
                SELDTMNO_RS.MoreResults
                SELDTMNO_RS.Requery
            End If
            If Trim(SELRREQUISITN_RS11("SER_CODE")) <> "8888" Then
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS11("ser_code"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQUISITN_RS11("ser_code"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If SELDREQFLAG = False Then
                    SELDREQQR(0) = (SELRREQUISITN_RS11("reqnno"))
                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                    SELDREQFLAG = True
                Else
                    SELDREQQR(0) = (SELRREQUISITN_RS11("reqnno"))
                    SELDREQRS.MoreResults
                    SELDREQRS.Requery
                End If
                If SELMA101_RS.RowCount > 0 Then
                    If Trim(SELRREQUISITN_RS11("REQNSTAT")) = "E" Then
                        If SELDREQRS.RowCount > 0 Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) & vbTab & Trim(SELDREQRS("PATHNO"))
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn"))
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn"))
                    End If
                    i = i + 1
                End If
            Else
                If SELDREQFLAG = False Then
                    SELDREQQR(0) = (SELRREQUISITN_RS11("reqnno"))
                    Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                    SELDREQFLAG = True
                Else
                    SELDREQQR(0) = (SELRREQUISITN_RS11("reqnno"))
                    SELDREQRS.MoreResults
                    SELDREQRS.Requery
                End If
                If Trim(SELRREQUISITN_RS11("REQNSTAT")) = "E" Then
                    If SELDREQRS.RowCount > 0 Then
                        If Not SELDTMNO_RS.EOF Then
                            If IsNull(SELDTMNO_RS("dtmno")) Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            End If
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                        End If
                    Else
                        If Not SELDTMNO_RS.EOF Then
                            If IsNull(SELDTMNO_RS("dtmno")) Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            End If
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                        End If
                    End If
                Else
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("dtmno")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" '& vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQUISITN_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                        End If
                    Else
                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQUISITN_RS11("reqnno")) & vbTab & SELRREQUISITN_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" '& vbTab & Trim(SELRREQUISITN_RS11("examinatn")) '" " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                    End If
                End If
                i = i + 1
            End If
            SELRREQUISITN_RS11.MoveNext
        Wend
        'chk_rcvd.Visible = False
        txt_pathno.Visible = False
    End If
End If
End Sub

Public Sub DTM1()
If SELDTMNO_FLAG = False Then
    SELDTMNO_QR(0) = msk_caseno.FormattedText
    Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
    SELDTMNO_FLAG = True
Else
    SELDTMNO_QR(0) = msk_caseno.FormattedText
    SELDTMNO_RS.MoreResults
    SELDTMNO_RS.Requery
End If
If NEWFLAG = True Then
    'NEWFLAG = False
    If SELRREQNOFLAG1 = False Then
        SELRREQNO_QR1(0) = msk_caseno.FormattedText
        SELRREQNO_QR1(1) = "C"
        SELRREQNO_QR1(2) = "C"
        SELRREQNO_QR1(3) = "P"   ' Made P from C on 27-01-2004
        SELRREQNO_QR1(4) = "C"
        Set SELRREQNO_RS1 = SELRREQNO_QR1.OpenResultset(1, 3)
        SELRREQNOFLAG1 = True
    Else
        SELRREQNO_QR1(0) = msk_caseno.FormattedText
        SELRREQNO_QR1(1) = "C"
        SELRREQNO_QR1(2) = "C"
        SELRREQNO_QR1(3) = "P"  ' Made P from C on 27-01-2004
        SELRREQNO_QR1(4) = "C"
        Set SELRREQNO_RS1 = SELRREQNO_QR1.OpenResultset(1, 3)
        SELRREQNO_RS1.MoreResults
        SELRREQNO_RS1.Requery
    End If
    If SELRREQNO_RS1.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQNO_RS1.EOF
            If SELRREQNO_RS1("SER_CODE") <> "8888" Then
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQNO_RS1("ser_code"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQNO_RS1("ser_code"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If Trim(MODALITY) = "SP" Then
                    If SELDREQFLAG = False Then
                        SELDREQQR(0) = (SELRREQNO_RS1("reqnno"))
                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                        SELDREQFLAG = True
                    Else
                        SELDREQQR(0) = (SELRREQNO_RS1("reqnno"))
                        SELDREQRS.MoreResults
                        SELDREQRS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQNO_RS1("reqnstat")) = "C" Then
                            If SELDREQRS.RowCount > 0 Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & SELDREQRS("PATHNO")
                                'chk_collected.Value = 1
                                'txt_pathno.Visible = False
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & ""
                                'chk_collected.Value = 1
                                'txt_pathno.Visible = False
                            End If
                        Else
                            If SELDREQRS.RowCount > 0 Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & Trim(SELDREQRS("PATHNO"))
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & ""
                            End If
                        End If
                        i = i + 1
                    End If
                Else
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQNO_RS1("reqnstat")) = "C" Then
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & ""
                                    GoTo n
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & "" & Trim(SELDTMNO_RS("DTMNO"))
                                    GoTo n
                                End If
                            End If
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                        Else
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " "
                                    GoTo n
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                    GoTo n
                                End If
                            End If
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " "
                        End If
n:
                        i = i + 1
                    End If
                End If
            Else
                If Trim(SELRREQNO_RS1("reqnstat")) = "C" Then
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & ""
                            GoTo m
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & "" & Trim(SELDTMNO_RS("DTMNO"))
                            GoTo m
                        End If
                    End If
                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & ""
                Else
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " "
                            GoTo m
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            GoTo m
                        End If
                    End If
                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " "
                End If
m:
                i = i + 1
            End If
            SELRREQNO_RS1.MoveNext
        Wend
        If Trim(MODALITY) = "SP" Then
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
        Else
            chk_rcvd.Visible = False
            txt_pathno.Visible = True
        End If
        If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or Trim(MODALITY) = "ALL" Then
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 5
            chk_rcvd.Width = MSFlexGrid1.CellWidth
            chk_rcvd.Height = MSFlexGrid1.CellHeight
            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_rcvd.Value = 0
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
        End If
        btnenabl ("NEW")
        'chk_collected.Visible = True
    Else
        MsgBox "No pending Requisition No."
        'chk_rcvd.Visible = False
        txt_pathno.Visible = False
        'btnenabl ("INI")
        FLDENABL ("DIS")
        'cmd_new.SetFocus
        Exit Sub
    End If
End If
If viewflag = True Then
'viewflag = False
If SELRREQNOFLAG1 = False Then
        SELRREQNO_QR1(0) = msk_caseno.FormattedText
        SELRREQNO_QR1(1) = "E"
        SELRREQNO_QR1(2) = "E"
        SELRREQNO_QR1(3) = "T"
        SELRREQNO_QR1(4) = "F"
        Set SELRREQNO_RS1 = SELRREQNO_QR1.OpenResultset(1, 3)
        SELRREQNOFLAG1 = True
    Else
        SELRREQNO_QR1(0) = msk_caseno.FormattedText
        SELRREQNO_QR1(1) = "E"
        SELRREQNO_QR1(2) = "E"
        SELRREQNO_QR1(3) = "T"
        SELRREQNO_QR1(4) = "F"
        Set SELRREQNO_RS1 = SELRREQNO_QR1.OpenResultset(1, 3)
        SELRREQNO_RS1.MoreResults
        SELRREQNO_RS1.Requery
    End If
        If SELRREQNO_RS1.RowCount > 0 Then
            i = 1
            MSFlexGrid1.Rows = 1
            While Not SELRREQNO_RS1.EOF
                If SELRREQNO_RS1("SER_CODE") <> "8888" Then
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = Trim(SELRREQNO_RS1("ser_code"))
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = Trim(SELRREQNO_RS1("ser_code"))
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            If SELDREQFLAG = False Then
                                SELDREQQR(0) = (SELRREQNO_RS1("reqnno"))
                                Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                SELDREQFLAG = True
                            Else
                                SELDREQQR(0) = (SELRREQNO_RS1("reqnno"))
                                SELDREQRS.MoreResults
                                SELDREQRS.Requery
                            End If
                            If Trim(SELRREQNO_RS1("REQNSTAT")) = "E" Then
                                If SELDREQRS.RowCount > 0 Then
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                                            GoTo m1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                Else
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                End If
                            Else
                                If Trim(SELRREQNO_RS1("REQNSTAT")) = "T" Or Trim(SELRREQNO_RS1("REQNSTAT")) = "T" Then
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                                            GoTo m1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                                End If
                            End If
m1:
                        i = i + 1
                        End If
                Else
                    If SELDREQFLAG = False Then
                        SELDREQQR(0) = (SELRREQNO_RS1("reqnno"))
                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                        SELDREQFLAG = True
                    Else
                        SELDREQQR(0) = (SELRREQNO_RS1("reqnno"))
                        SELDREQRS.MoreResults
                        SELDREQRS.Requery
                    End If
                    If Trim(SELRREQNO_RS1("REQNSTAT")) = "E" Then
                            If SELDREQRS.RowCount > 0 Then
                                If Not SELDTMNO_RS.EOF Then
                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                                        GoTo N1
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                            Else
                                If Not SELDTMNO_RS.EOF Then
                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) ' & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) ' & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            End If
                        Else
                            If Trim(SELRREQNO_RS1("REQNSTAT")) = "T" Or Trim(SELRREQNO_RS1("REQNSTAT")) = "T" Then
                                If Not SELDTMNO_RS.EOF Then
                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                                        GoTo N1
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS1("reqnno")) & vbTab & SELRREQNO_RS1("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS1("examinatn"))
                            End If
                        End If
N1:
                     i = i + 1
                End If
                SELRREQNO_RS1.MoveNext
            Wend
           'chk_rcvd.Visible = False
           txt_pathno.Visible = False
           'chk_collected.Visible = False
        End If
    End If
'If opt_staff.Value = True Then
'    STAFF
'End If

End Sub
Public Sub DTM1MODWISE()
If SELDTMNO_FLAG = False Then
    SELDTMNO_QR(0) = msk_caseno.FormattedText
    Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
    SELDTMNO_FLAG = True
Else
    SELDTMNO_QR(0) = msk_caseno.FormattedText
    SELDTMNO_RS.MoreResults
    SELDTMNO_RS.Requery
End If
If NEWFLAG = True Then
    'NEWFLAG = False
    If SELRREQNO_FLAG11 = False Then
        SELRREQNO_QR11(0) = MODALITY
        SELRREQNO_QR11(1) = msk_caseno.FormattedText
        SELRREQNO_QR11(2) = "C"
        SELRREQNO_QR11(3) = "P"   ' Made P from C on 27-01-2004
        SELRREQNO_QR11(4) = "C"
        'SELRREQNO_QR1(4) = "P"
        Set SELRREQNO_RS11 = SELRREQNO_QR11.OpenResultset(1, 3)
        SELRREQNO_FLAG11 = True
    Else
        SELRREQNO_QR11(0) = Trim(MODALITY)
        SELRREQNO_QR11(1) = msk_caseno.FormattedText
        SELRREQNO_QR11(2) = "C"
        SELRREQNO_QR11(3) = "P"   ' Made P from C on 27-01-2004
        SELRREQNO_QR11(4) = "C"
        'SELRREQNO_QR1(4) = "P"
        Set SELRREQNO_RS11 = SELRREQNO_QR11.OpenResultset(1, 3)
        SELRREQNO_RS11.MoreResults
        SELRREQNO_RS11.Requery
    End If
    If SELRREQNO_RS11.RowCount > 0 Then
        i = 1
        MSFlexGrid1.Rows = 1
        While Not SELRREQNO_RS11.EOF
            If SELRREQNO_RS11("SER_CODE") <> "8888" Then
                If SELMA101FLAG = False Then
                    SELMA101_QR(0) = Trim(SELRREQNO_RS11("ser_code"))
                    Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                    SELMA101FLAG = True
                Else
                    SELMA101_QR(0) = Trim(SELRREQNO_RS11("ser_code"))
                    SELMA101_RS.MoreResults
                    SELMA101_RS.Requery
                End If
                If Trim(MODALITY) = "SP" Then
                    If SELDREQFLAG = False Then
                        SELDREQQR(0) = (SELRREQNO_RS11("reqnno"))
                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                        SELDREQFLAG = True
                    Else
                        SELDREQQR(0) = (SELRREQNO_RS11("reqnno"))
                        SELDREQRS.MoreResults
                        SELDREQRS.Requery
                    End If
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQNO_RS11("reqnstat")) = "C" Then
                            If SELDREQRS.RowCount > 0 Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & SELDREQRS("PATHNO")
                                'chk_collected.Value = 1
                                'txt_pathno.Visible = False
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & ""
                                'chk_collected.Value = 1
                                'txt_pathno.Visible = False
                            End If
                        Else
                            If SELDREQRS.RowCount > 0 Then
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & Trim(SELDREQRS("PATHNO"))
                            Else
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & ""
                            End If
                        End If
                        i = i + 1
                    End If
                Else
                    If SELMA101_RS.RowCount > 0 Then
                        If Trim(SELRREQNO_RS11("reqnstat")) = "C" Then
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & ""
                                    GoTo n
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & " " & vbTab & "" & Trim(SELDTMNO_RS("DTMNO"))
                                    GoTo n
                                End If
                            End If
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED"
                        Else
                            If Not SELDTMNO_RS.EOF Then
                                If IsNull(SELDTMNO_RS("DTMNO")) Then
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " "
                                    GoTo n
                                Else
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                    GoTo n
                                End If
                            End If
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & " " & vbTab & " "
                        End If
n:
                        i = i + 1
                    End If
                End If
            Else
                If Trim(SELRREQNO_RS11("reqnstat")) = "C" Then
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & ""
                            GoTo m
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & "" & Trim(SELDTMNO_RS("DTMNO"))
                            GoTo m
                        End If
                    End If
                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED " & vbTab & "" & vbTab & ""
                Else
                    If Not SELDTMNO_RS.EOF Then
                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " "
                            GoTo m
                        Else
                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " " & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            GoTo m
                        End If
                    End If
                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & " " & vbTab & " "
                End If
m:
                i = i + 1
            End If
            SELRREQNO_RS11.MoveNext
        Wend

''''        If Trim(MODALITY) = "SP" Then
''''            chk_rcvd.Visible = True
''''            txt_pathno.Visible = False
''''        Else
''''            chk_rcvd.Visible = False
''''            txt_pathno.Visible = True
''''        End If
        If Trim(MODALITY) = "SP" Then
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
            chk_rcvd.SetFocus
        Else
            chk_rcvd.Visible = False
            txt_pathno.Visible = True
        End If
        'chk_collected.Visible = True
        If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or Trim(MODALITY) = "ALL" Then
            MSFlexGrid1.row = 1
            MSFlexGrid1.col = 5
            chk_rcvd.Width = MSFlexGrid1.CellWidth
            chk_rcvd.Height = MSFlexGrid1.CellHeight
            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            chk_rcvd.Visible = True
            txt_pathno.Visible = False
        End If


''        If Trim(MODALITY) = "BP" Or Trim(MODALITY) = "SS" Or Trim(MODALITY) = "HT" Or Trim(MODALITY) = "ALL" Then
''            MSFlexGrid1.Row = 1
''            MSFlexGrid1.Col = 5
''            chk_rcvd.Width = MSFlexGrid1.CellWidth
''            chk_rcvd.Height = MSFlexGrid1.CellHeight
''            chk_rcvd.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
''            chk_rcvd.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
''            chk_rcvd.Value = 0
''            chk_rcvd.Visible = True
''            txt_pathno.Visible = False
''        End If
        btnenabl ("NEW")
        'chk_collected.Visible = True
    Else
        MsgBox "No pending Requisition No."
        'chk_rcvd.Visible = False
        txt_pathno.Visible = False
        'btnenabl ("INI")
        FLDENABL ("DIS")
        'cmd_new.SetFocus
        Exit Sub
    End If
End If
If viewflag = True Then
'viewflag = False
If SELRREQNO_FLAG11 = False Then
        SELRREQNO_QR11(0) = MODALITY
        SELRREQNO_QR11(1) = msk_caseno.FormattedText
        SELRREQNO_QR11(2) = "E"
        SELRREQNO_QR11(3) = "T"
        SELRREQNO_QR11(4) = "E"
        'SELRREQNO_QR1(4) = "F"
        Set SELRREQNO_RS11 = SELRREQNO_QR11.OpenResultset(1, 3)
        SELRREQNOFLAG1 = True
    Else
        SELRREQNO_QR11(0) = MODALITY
        SELRREQNO_QR11(1) = msk_caseno.FormattedText
        SELRREQNO_QR11(2) = "E"
        SELRREQNO_QR11(3) = "T"
        SELRREQNO_QR11(4) = "F"
        'SELRREQNO_QR1(4) = "F"
        Set SELRREQNO_RS11 = SELRREQNO_QR11.OpenResultset(1, 3)
        SELRREQNO_RS11.MoreResults
        SELRREQNO_RS11.Requery
    End If
        If SELRREQNO_RS11.RowCount > 0 Then
            i = 1
            MSFlexGrid1.Rows = 1
            While Not SELRREQNO_RS11.EOF
                If SELRREQNO_RS11("SER_CODE") <> "8888" Then
                        If SELMA101FLAG = False Then
                            SELMA101_QR(0) = Trim(SELRREQNO_RS11("ser_code"))
                            Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
                            SELMA101FLAG = True
                        Else
                            SELMA101_QR(0) = Trim(SELRREQNO_RS11("ser_code"))
                            SELMA101_RS.MoreResults
                            SELMA101_RS.Requery
                        End If
                        If SELMA101_RS.RowCount > 0 Then
                            If SELDREQFLAG = False Then
                                SELDREQQR(0) = (SELRREQNO_RS11("reqnno"))
                                Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                                SELDREQFLAG = True
                            Else
                                SELDREQQR(0) = (SELRREQNO_RS11("reqnno"))
                                SELDREQRS.MoreResults
                                SELDREQRS.Requery
                            End If
                            If Trim(SELRREQNO_RS11("REQNSTAT")) = "E" Then
                                If SELDREQRS.RowCount > 0 Then
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                                            GoTo m1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                Else
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) '& vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                End If
                            Else
                                If Trim(SELRREQNO_RS11("REQNSTAT")) = "T" Or Trim(SELRREQNO_RS11("REQNSTAT")) = "T" Then
                                    If Not SELDTMNO_RS.EOF Then
                                        If IsNull(SELDTMNO_RS("DTMNO")) Then
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                                            GoTo m1
                                        Else
                                            MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                            GoTo m1
                                        End If
                                    End If
                                    MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & Trim(SELMA101_RS("ser_det")) & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                                End If
                            End If
m1:
                        i = i + 1
                        End If
                Else
                    If SELDREQFLAG = False Then
                        SELDREQQR(0) = (SELRREQNO_RS11("reqnno"))
                        Set SELDREQRS = SELDREQQR.OpenResultset(1, 3)
                        SELDREQFLAG = True
                    Else
                        SELDREQQR(0) = (SELRREQNO_RS11("reqnno"))
                        SELDREQRS.MoreResults
                        SELDREQRS.Requery
                    End If
                    If Trim(SELRREQNO_RS11("REQNSTAT")) = "E" Then
                            If SELDREQRS.RowCount > 0 Then
                                If Not SELDTMNO_RS.EOF Then
                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                                        GoTo N1
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                            Else
                                If Not SELDTMNO_RS.EOF Then
                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) ' & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) ' & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                            End If
                        Else
                            If Trim(SELRREQNO_RS11("REQNSTAT")) = "T" Or Trim(SELRREQNO_RS11("REQNSTAT")) = "T" Then
                                If Not SELDTMNO_RS.EOF Then
                                    If IsNull(SELDTMNO_RS("DTMNO")) Then
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                                        GoTo N1
                                    Else
                                        MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn")) & vbTab & Trim(SELDTMNO_RS("DTMNO"))
                                        GoTo N1
                                    End If
                                End If
                                MSFlexGrid1.AddItem i & vbTab & req_chglen(SELRREQNO_RS11("reqnno")) & vbTab & SELRREQNO_RS11("caseno") & vbTab & "M.SERVICE" & vbTab & "COLLECTED" & vbTab & Trim(SELRREQNO_RS11("examinatn"))
                            End If
                        End If
N1:
                     i = i + 1
                End If
                SELRREQNO_RS11.MoveNext
            Wend
           'chk_rcvd.Visible = False
           txt_pathno.Visible = False
           'chk_collected.Visible = False
        End If
    End If
'If opt_staff.Value = True Then
    'STAFF
'End If

End Sub

Public Sub SPAUTONO()
Dim PATH As String
On Error GoTo MSG:
i = 1
For i = 1 To MSFlexGrid1.Rows - 1
    If Trim(MSFlexGrid1.TextMatrix(i, 5)) <> "" Then
        If SELDAUTOPATHFLAG = False Then
            SELDAUTOPATH_QR(0) = Trim(dept)
            SELDAUTOPATH_QR(1) = Trim(MODALITY)
            SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
            Set SELDAUTOPATH_RS = SELDAUTOPATH_QR.OpenResultset(2, 3)
            SELDAUTOPATHFLAG = True
        Else
            SELDAUTOPATH_QR(0) = Trim(dept)
            SELDAUTOPATH_QR(1) = Trim(MODALITY)
            SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
            SELDAUTOPATH_RS.MoreResults
            SELDAUTOPATH_RS.Requery
        End If
        If SELDAUTOPATH_RS.RowCount > 0 Then
            AUTOPATHNO = SELDAUTOPATH_RS("PATHNO") + 1
            UPDDAUTOPATH_QR(0) = AUTOPATHNO
            UPDDAUTOPATH_QR(1) = Trim(dept)
            UPDDAUTOPATH_QR(2) = Trim(MODALITY)
            UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            UPDDAUTOPATH_QR.Execute
        Else
            AUTOPATHNO = 1
            INSDAUTOPATH_QR(0) = Trim(dept)
            INSDAUTOPATH_QR(1) = Trim(MODALITY)
            INSDAUTOPATH_QR(2) = Trim(Left(msk_caseno.FormattedText, 2))
            INSDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            INSDAUTOPATH_QR(4) = 1
            INSDAUTOPATH_QR.Execute
        End If
        If Trim(dept) = "UZZ" Then
                PATH = Trim(SELDAUTOPATH_RS("PATHYEAR"))
        Else
            If DPATHYEARFLAG = False Then
                DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
                Set DPATHYEAR_RS = DPATHYEAR_QR.OpenResultset(1, 3)
                DPATHYEARFLAG = True
            Else
                DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
                DPATHYEAR_RS.MoreResults
                DPATHYEAR_RS.Requery
            End If
            If DPATHYEAR_RS.RowCount > 0 Then
                PATH = DPATHYEAR_RS("LASTTWO")
            End If
        End If
    '***********padding the reqslno with leading zeros"****************
        Dim reqslno1
        reqslno1 = AUTOPATHNO
        If Trim(dept) = "UZZ" Then
            While Len(CStr(reqslno1)) < 6  ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
            ' NEEDED NUMBERING LIKE
            ' F05999999 FOR BF
            ' P05999999 FOR HL
            ' M05999999 FOR BM

                reqslno1 = "0" & reqslno1
            Wend
        Else
            While Len(CStr(reqslno1)) < 6
                reqslno1 = "0" & reqslno1
            Wend
        End If
        DAUTOPATHNO = reqslno1
'******************************************************************
        'If Trim(MODALITY) = "SP" Then
        If Trim(dept) = "UZZ" Then
            ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
            ' NEEDED NUMBERING LIKE
            ' F05999999 FOR BF
            ' P05999999 FOR HL
            ' M05999999 FOR BM
            MSFlexGrid1.TextMatrix(i, 6) = Trim(PATH) & CStr(Trim(Right(DATE11, 2))) & Trim(DAUTOPATHNO)
        Else
            MSFlexGrid1.TextMatrix(i, 6) = DAUTOPATHNO & "/" & PATH
        End If
        'End If
    End If
Next i
Exit Sub
MSG:
MsgBox Err.Description
End Sub

Public Sub TMAUTONO()
 i = 1
For i = 1 To MSFlexGrid1.Rows - 1
    If Trim(MSFlexGrid1.TextMatrix(i, 5)) <> "" Then
        SELTMAUTOPATH_QR(0) = Trim(dept)
        SELTMAUTOPATH_QR(1) = Trim(MODALITY)
        SELTMAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
        SELTMAUTOPATH_QR(3) = Trim(Mid(DATE11, 4, 2))
        SELTMAUTOPATH_QR(4) = Trim(Left(DATE11, 2))
        If SELTMAUTOPATHFLAG = False Then
            Set SELTMAUTOPATH_RS = SELTMAUTOPATH_QR.OpenResultset(2, 3)
            SELTMAUTOPATHFLAG = True
        Else
            SELTMAUTOPATH_RS.MoreResults
            SELTMAUTOPATH_RS.Requery
        End If
        If SELTMAUTOPATH_RS.RowCount > 0 Then
            AUTOPATHNO = SELTMAUTOPATH_RS("PATHNO") + 1
            UPDTMAUTOPATH_QR(0) = AUTOPATHNO
            UPDTMAUTOPATH_QR(1) = Trim(dept)
            UPDTMAUTOPATH_QR(2) = Trim(MODALITY)
            UPDTMAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            UPDTMAUTOPATH_QR(4) = Trim(Mid(DATE11, 4, 2))
            UPDTMAUTOPATH_QR(5) = Trim(Left(DATE11, 2))
            UPDTMAUTOPATH_QR.Execute
        Else
            AUTOPATHNO = 1
            INSTMAUTOPATH_QR(0) = Trim(dept)
            INSTMAUTOPATH_QR(1) = Trim(MODALITY)
            INSTMAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
            INSTMAUTOPATH_QR(3) = Trim(Mid(DATE11, 4, 2))
            INSTMAUTOPATH_QR(4) = Trim(Left(DATE11, 2))
            INSTMAUTOPATH_QR(5) = 1
            INSTMAUTOPATH_QR.Execute
        End If

        Dim reqslno1
        reqslno1 = AUTOPATHNO
        If Trim(dept) = "FEC" Then
            While Len(CStr(reqslno1)) < 3
                reqslno1 = "0" & reqslno1
            Wend
        End If
        DAUTOPATHNO = reqslno1
        MSFlexGrid1.TextMatrix(i, 6) = Trim(Right(DATE11, 2)) & Trim(Mid(DATE11, 4, 2)) & Trim(Left(DATE11, 2)) & DAUTOPATHNO
    End If
Next i
Exit Sub
MSG:
MsgBox Err.Description

End Sub

'ADDED BY MILIND ON 25-05-2006 AS NEW BAR CODE
Public Sub PRINTLBLNEW(BAR_CD1 As String)

BAR_CD1 = Trim(BAR_CD1)
bar1 = Mid(BAR_CD1, 1, 2)
bar2 = Mid(BAR_CD1, 3, 8)
bar3 = Mid(BAR_CD1, 11, 1)
f1 = 1
Print #f1, "{D0281,0500,0251|}"
Print #f1, "{AY;+10,0|}"
Print #f1, "{C|}"
Print #f1, "{PC000;0042,0025,05,05,K,00,B|}"
Print #f1, "{PC001;0403,0025,05,05,K,00,B|}"
Print #f1, "{PC002;0163,0025,05,05,K,00,B|}"
Print #f1, "{PC003;0042,0060,07,07,I,00,B|}"
Print #f1, "{PC004;0133,0060,07,07,I,00,B|}"
Print #f1, "{PC005;0223,0060,07,07,I,00,B|}"
Print #f1, "{PC006;0313,0060,07,07,I,00,B|}"
Print #f1, "{PC007;0403,0060,07,07,I,00,B|}"
Print #f1, "{PC008;0042,0091,07,07,I,00,B|}"
Print #f1, "{PC009;0133,0091,07,07,I,00,B|}"
Print #f1, "{PC010;0223,0091,07,07,I,00,B|}"
Print #f1, "{PC011;0313,0091,07,07,I,00,B|}"
Print #f1, "{PC012;0403,0091,07,07,I,00,B|}"
Print #f1, "{PC013;0042,0120,07,07,I,00,B|}"
Print #f1, "{XB00;0133,0110,9,3,02,0,0065,+0000000000,000,0,00|}"
Print #f1, "{PC014;0153,0200,05,05,I,00,B|}"
Print #f1, "{PC015;0042,0231,05,05,I,00,B|}"
Print #f1, "{PC016;0403,0231,05,05,I,00,B|}"
Print #f1, "{RC000;TMH|}"
Print #f1, "{RC001;" + LABNO + "|}"
Print #f1, "{RC002;" + CASENO + "|}"

If BARCODE_NEW = True Then
    Dim i, c, PC(11) As Variant
    Dim RC As Integer
    SELPROCODE_QR(0) = REQNO
    SELPROCODE_QR.Execute
    Set SELPROCODE_RS = SELPROCODE_QR.OpenResultset(1, 3)
    c = SELPROCODE_RS.RowCount

    Dim flg1, flg2, flg3, flg4, flg5, flg6, flg7, flg8, flg9, flg10, flg1a As Boolean
    i = 1
code:   While SELPROCODE_RS.EOF = False
            PC(i) = SELPROCODE_RS("TESTCODE")

            If c = 1 Then
                Print #f1, "{RC003;" + PC(i) + "|}"
            End If

            If c = 2 Then
                If flg1 = False Then
                    Print #f1, "{RC003;" + PC(i) + "|}"
                End If

                    flg1 = True

                If flg2 = True Then
                    Print #f1, "{RC004;" + PC(i) + "|}"
                End If

                    flg2 = True
                    SELPROCODE_RS.MoveNext
                    i = i + 1
                    GoTo code

             End If

             If c = 3 Then
                If flg1 = False Then
                    Print #f1, "{RC003;" + PC(i) + "|}"
                End If

                    flg1 = True

                If flg2 = True Then
                    If i = 2 Then
                        Print #f1, "{RC004;" + PC(i) + "|}"
                    End If
                End If

                    flg2 = True

                If i = 3 Then
                    flg3 = True
                    If flg3 = True Then
                        Print #f1, "{RC005;" + PC(i) + "|}"
                    End If
                End If

                    SELPROCODE_RS.MoveNext
                    i = i + 1
                    GoTo code
              End If

              If c = 4 Then
                     If flg1 = False Then
                         Print #f1, "{RC003;" + PC(i) + "|}"
                     End If

                         flg1 = True

                     If flg2 = True Then
                         If i = 2 Then
                             Print #f1, "{RC004;" + PC(i) + "|}"
                         End If
                     End If

                         flg2 = True

                     If i = 3 Then
                         flg3 = True
                         If flg3 = True Then
                             Print #f1, "{RC005;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 4 Then
                         flg4 = True
                         If flg4 = True Then
                             Print #f1, "{RC006;" + PC(i) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         i = i + 1
                         GoTo code
                End If


                If c = 6 Then
                     If flg1 = False Then
                         Print #f1, "{RC003;" + PC(i) + "|}"
                     End If

                         flg1 = True

                     If flg2 = True Then
                         If i = 2 Then
                             Print #f1, "{RC004;" + PC(i) + "|}"
                         End If
                     End If

                         flg2 = True

                     If i = 3 Then
                         flg3 = True
                         If flg3 = True Then
                             Print #f1, "{RC005;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 4 Then
                         flg4 = True
                         If flg4 = True Then
                             Print #f1, "{RC006;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 5 Then
                         flg5 = True
                         If flg5 = True Then
                             Print #f1, "{RC007;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 6 Then
                         flg6 = True
                         If flg6 = True Then
                             Print #f1, "{RC008;" + PC(i) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         i = i + 1
                         GoTo code
                End If

                If c = 7 Then
                     If flg1 = False Then
                         Print #f1, "{RC003;" + PC(i) + "|}"
                     End If

                         flg1 = True

                     If flg2 = True Then
                         If i = 2 Then
                             Print #f1, "{RC004;" + PC(i) + "|}"
                         End If
                     End If

                         flg2 = True

                     If i = 3 Then
                         flg3 = True
                         If flg3 = True Then
                             Print #f1, "{RC005;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 4 Then
                         flg4 = True
                         If flg4 = True Then
                             Print #f1, "{RC006;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 5 Then
                         flg5 = True
                         If flg5 = True Then
                             Print #f1, "{RC007;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 6 Then
                         flg6 = True
                         If flg6 = True Then
                             Print #f1, "{RC008;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 7 Then
                         flg7 = True
                         If flg7 = True Then
                             Print #f1, "{RC009;" + PC(i) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         i = i + 1
                         GoTo code
                End If

                If c = 8 Then
                     If flg1 = False Then
                         Print #f1, "{RC003;" + PC(i) + "|}"
                     End If

                         flg1 = True

                     If flg2 = True Then
                         If i = 2 Then
                             Print #f1, "{RC004;" + PC(i) + "|}"
                         End If
                     End If

                         flg2 = True

                     If i = 3 Then
                         flg3 = True
                         If flg3 = True Then
                             Print #f1, "{RC005;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 4 Then
                         flg4 = True
                         If flg4 = True Then
                             Print #f1, "{RC006;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 5 Then
                         flg5 = True
                         If flg5 = True Then
                             Print #f1, "{RC007;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 6 Then
                         flg6 = True
                         If flg6 = True Then
                             Print #f1, "{RC008;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 7 Then
                         flg7 = True
                         If flg7 = True Then
                             Print #f1, "{RC009;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 8 Then
                         flg8 = True
                         If flg8 = True Then
                             Print #f1, "{RC010;" + PC(i) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         i = i + 1
                         GoTo code
                End If

                If c = 9 Then
                     If flg1 = False Then
                         Print #f1, "{RC003;" + PC(i) + "|}"
                     End If

                         flg1 = True

                     If flg2 = True Then
                         If i = 2 Then
                             Print #f1, "{RC004;" + PC(i) + "|}"
                         End If
                     End If

                         flg2 = True

                     If i = 3 Then
                         flg3 = True
                         If flg3 = True Then
                             Print #f1, "{RC005;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 4 Then
                         flg4 = True
                         If flg4 = True Then
                             Print #f1, "{RC006;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 5 Then
                         flg5 = True
                         If flg5 = True Then
                             Print #f1, "{RC007;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 6 Then
                         flg6 = True
                         If flg6 = True Then
                             Print #f1, "{RC008;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 7 Then
                         flg7 = True
                         If flg7 = True Then
                             Print #f1, "{RC009;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 8 Then
                         flg8 = True
                         If flg8 = True Then
                             Print #f1, "{RC010;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 9 Then
                         flg9 = True
                         If flg9 = True Then
                             Print #f1, "{RC011;" + PC(i) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         i = i + 1
                         GoTo code
                End If

                If c = 10 Then
                     If flg1 = False Then
                         Print #f1, "{RC003;" + PC(i) + "|}"
                     End If

                         flg1 = True

                     If flg2 = True Then
                         If i = 2 Then
                             Print #f1, "{RC004;" + PC(i) + "|}"
                         End If
                     End If

                         flg2 = True

                     If i = 3 Then
                         flg3 = True
                         If flg3 = True Then
                             Print #f1, "{RC005;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 4 Then
                         flg4 = True
                         If flg4 = True Then
                             Print #f1, "{RC006;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 5 Then
                         flg5 = True
                         If flg5 = True Then
                             Print #f1, "{RC007;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 6 Then
                         flg6 = True
                         If flg6 = True Then
                             Print #f1, "{RC008;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 7 Then
                         flg7 = True
                         If flg7 = True Then
                             Print #f1, "{RC009;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 8 Then
                         flg8 = True
                         If flg8 = True Then
                             Print #f1, "{RC010;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 9 Then
                         flg9 = True
                         If flg9 = True Then
                             Print #f1, "{RC011;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 10 Then
                         flg10 = True
                         If flg10 = True Then
                             Print #f1, "{RC012;" + PC(i) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         i = i + 1
                         GoTo code
                End If

                If c = 11 Then
                     If flg1 = False Then
                         Print #f1, "{RC003;" + PC(i) + "|}"
                     End If

                         flg1 = True

                     If flg2 = True Then
                         If i = 2 Then
                             Print #f1, "{RC004;" + PC(i) + "|}"
                         End If
                     End If

                         flg2 = True

                     If i = 3 Then
                         flg3 = True
                         If flg3 = True Then
                             Print #f1, "{RC005;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 4 Then
                         flg4 = True
                         If flg4 = True Then
                             Print #f1, "{RC006;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 5 Then
                         flg5 = True
                         If flg5 = True Then
                             Print #f1, "{RC007;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 6 Then
                         flg6 = True
                         If flg6 = True Then
                             Print #f1, "{RC008;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 7 Then
                         flg7 = True
                         If flg7 = True Then
                             Print #f1, "{RC009;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 8 Then
                         flg8 = True
                         If flg8 = True Then
                             Print #f1, "{RC010;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 9 Then
                         flg9 = True
                         If flg9 = True Then
                             Print #f1, "{RC011;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 10 Then
                         flg10 = True
                         If flg10 = True Then
                             Print #f1, "{RC012;" + PC(i) + "|}"
                         End If
                     End If

                     If i = 11 Then
                         flg1a = True
                         If flg1a = True Then
                             Print #f1, "{RC013;" + PC(i) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         i = i + 1
                         GoTo code
                End If
         Wend
         Print #f1, "{RB00;" + REQNO + "|}"
         Print #f1, "{RC014;" + REQNO + "|}"
         Print #f1, "{RC015;" + REC_DATE + "|}"
         Print #f1, "{RC016;" + REC_TIME + "|}"
         Print #f1, "{XS;I,0001,0002C4201|}"
End If

If BARCODE_VIEW = True Then
    Dim I1, C1, PC1(11) As Variant
    Dim RC1 As Integer
    SELPROCODE_QR(0) = II
    SELPROCODE_QR.Execute
    Set SELPROCODE_RS = SELPROCODE_QR.OpenResultset(1, 3)
    C1 = SELPROCODE_RS.RowCount

    Dim flg11, flg21, flg31, flg41, flg51, flg61, flg71, flg81, flg91, flg101, flg111 As Boolean
    I1 = 1
code1:   While SELPROCODE_RS.EOF = False
            PC1(I1) = SELPROCODE_RS("TESTCODE")

            If C1 = 1 Then
                Print #f1, "{RC003;" + PC1(I1) + "|}"
            End If

            If C1 = 2 Then
                If flg11 = False Then
                    Print #f1, "{RC003;" + PC1(I1) + "|}"
                End If

                    flg11 = True

                If flg21 = True Then
                    Print #f1, "{RC004;" + PC1(I1) + "|}"
                End If

                    flg21 = True
                    SELPROCODE_RS.MoveNext
                    I1 = I1 + 1
                    GoTo code1

             End If

             If C1 = 3 Then
                If flg11 = False Then
                    Print #f1, "{RC003;" + PC1(I1) + "|}"
                End If

                    flg11 = True

                If flg21 = True Then
                    If I1 = 2 Then
                        Print #f1, "{RC004;" + PC1(I1) + "|}"
                    End If
                End If

                    flg21 = True

                If I1 = 3 Then
                    flg31 = True
                    If flg31 = True Then
                        Print #f1, "{RC005;" + PC1(I1) + "|}"
                    End If
                End If

                    SELPROCODE_RS.MoveNext
                    I1 = I1 + 1
                    GoTo code1
              End If

              If C1 = 4 Then
                     If flg11 = False Then
                         Print #f1, "{RC003;" + PC1(I1) + "|}"
                     End If

                         flg11 = True

                     If flg21 = True Then
                         If I1 = 2 Then
                             Print #f1, "{RC004;" + PC1(I1) + "|}"
                         End If
                     End If

                         flg21 = True

                     If I1 = 3 Then
                         flg31 = True
                         If flg31 = True Then
                             Print #f1, "{RC005;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 4 Then
                         flg41 = True
                         If flg41 = True Then
                             Print #f1, "{RC006;" + PC1(I1) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         I1 = I1 + 1
                         GoTo code1
                End If


                If C1 = 6 Then
                     If flg11 = False Then
                         Print #f1, "{RC003;" + PC1(I1) + "|}"
                     End If

                         flg11 = True

                     If flg21 = True Then
                         If I1 = 2 Then
                             Print #f1, "{RC004;" + PC1(I1) + "|}"
                         End If
                     End If

                         flg2 = True

                     If I1 = 3 Then
                         flg31 = True
                         If flg31 = True Then
                             Print #f1, "{RC005;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 4 Then
                         flg41 = True
                         If flg41 = True Then
                             Print #f1, "{RC006;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 5 Then
                         flg51 = True
                         If flg51 = True Then
                             Print #f1, "{RC007;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 6 Then
                         flg61 = True
                         If flg61 = True Then
                             Print #f1, "{RC008;" + PC1(I1) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         I1 = I1 + 1
                         GoTo code1
                End If

                If C1 = 7 Then
                     If flg11 = False Then
                         Print #f1, "{RC003;" + PC1(I1) + "|}"
                     End If

                         flg11 = True

                     If flg21 = True Then
                         If I1 = 2 Then
                             Print #f1, "{RC004;" + PC1(I1) + "|}"
                         End If
                     End If

                         flg21 = True

                     If I1 = 3 Then
                         flg31 = True
                         If flg31 = True Then
                             Print #f1, "{RC005;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 4 Then
                         flg41 = True
                         If flg41 = True Then
                             Print #f1, "{RC006;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 5 Then
                         flg51 = True
                         If flg51 = True Then
                             Print #f1, "{RC007;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 6 Then
                         flg61 = True
                         If flg61 = True Then
                             Print #f1, "{RC008;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 7 Then
                         flg71 = True
                         If flg71 = True Then
                             Print #f1, "{RC009;" + PC1(I1) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         I1 = I1 + 1
                         GoTo code1
                End If

                If C1 = 8 Then
                     If flg11 = False Then
                         Print #f1, "{RC003;" + PC1(I1) + "|}"
                     End If

                         flg11 = True

                     If flg21 = True Then
                         If I1 = 2 Then
                             Print #f1, "{RC004;" + PC1(I1) + "|}"
                         End If
                     End If

                         flg21 = True

                     If I1 = 3 Then
                         flg31 = True
                         If flg31 = True Then
                             Print #f1, "{RC005;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 4 Then
                         flg41 = True
                         If flg41 = True Then
                             Print #f1, "{RC006;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 5 Then
                         flg51 = True
                         If flg51 = True Then
                             Print #f1, "{RC007;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 6 Then
                         flg61 = True
                         If flg61 = True Then
                             Print #f1, "{RC008;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 7 Then
                         flg71 = True
                         If flg71 = True Then
                             Print #f1, "{RC009;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 8 Then
                         flg81 = True
                         If flg81 = True Then
                             Print #f1, "{RC010;" + PC1(I1) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         I1 = I1 + 1
                         GoTo code1
                End If

                If C1 = 9 Then
                     If flg11 = False Then
                         Print #f1, "{RC003;" + PC1(I1) + "|}"
                     End If

                         flg11 = True

                     If flg21 = True Then
                         If I1 = 2 Then
                             Print #f1, "{RC004;" + PC1(I1) + "|}"
                         End If
                     End If

                         flg21 = True

                     If I1 = 3 Then
                         flg31 = True
                         If flg31 = True Then
                             Print #f1, "{RC005;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 4 Then
                         flg41 = True
                         If flg41 = True Then
                             Print #f1, "{RC006;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 5 Then
                         flg51 = True
                         If flg51 = True Then
                             Print #f1, "{RC007;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 6 Then
                         flg61 = True
                         If flg61 = True Then
                             Print #f1, "{RC008;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 7 Then
                         flg71 = True
                         If flg71 = True Then
                             Print #f1, "{RC009;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 8 Then
                         flg81 = True
                         If flg81 = True Then
                             Print #f1, "{RC010;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 9 Then
                         flg91 = True
                         If flg91 = True Then
                             Print #f1, "{RC011;" + PC1(I1) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         I1 = I1 + 1
                         GoTo code1
                End If

                If C1 = 10 Then
                     If flg11 = False Then
                         Print #f1, "{RC003;" + PC1(I1) + "|}"
                     End If

                         flg11 = True

                     If flg21 = True Then
                         If I1 = 2 Then
                             Print #f1, "{RC004;" + PC1(I1) + "|}"
                         End If
                     End If

                         flg21 = True

                     If I1 = 3 Then
                         flg31 = True
                         If flg31 = True Then
                             Print #f1, "{RC005;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 4 Then
                         flg41 = True
                         If flg41 = True Then
                             Print #f1, "{RC006;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 5 Then
                         flg51 = True
                         If flg51 = True Then
                             Print #f1, "{RC007;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 6 Then
                         flg61 = True
                         If flg61 = True Then
                             Print #f1, "{RC008;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 7 Then
                         flg71 = True
                         If flg71 = True Then
                             Print #f1, "{RC009;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 8 Then
                         flg81 = True
                         If flg81 = True Then
                             Print #f1, "{RC010;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 9 Then
                         flg91 = True
                         If flg91 = True Then
                             Print #f1, "{RC011;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 10 Then
                         flg101 = True
                         If flg101 = True Then
                             Print #f1, "{RC012;" + PC1(I1) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         I1 = I1 + 1
                         GoTo code1
                End If

                If C1 = 11 Then
                     If flg11 = False Then
                         Print #f1, "{RC003;" + PC1(I1) + "|}"
                     End If

                         flg11 = True

                     If flg21 = True Then
                         If I1 = 2 Then
                             Print #f1, "{RC004;" + PC1(I1) + "|}"
                         End If
                     End If

                         flg21 = True

                     If I1 = 3 Then
                         flg31 = True
                         If flg31 = True Then
                             Print #f1, "{RC005;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 4 Then
                         flg41 = True
                         If flg41 = True Then
                             Print #f1, "{RC006;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 5 Then
                         flg51 = True
                         If flg51 = True Then
                             Print #f1, "{RC007;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 6 Then
                         flg61 = True
                         If flg61 = True Then
                             Print #f1, "{RC008;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 7 Then
                         flg71 = True
                         If flg71 = True Then
                             Print #f1, "{RC009;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 8 Then
                         flg81 = True
                         If flg81 = True Then
                             Print #f1, "{RC010;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 9 Then
                         flg91 = True
                         If flg91 = True Then
                             Print #f1, "{RC011;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 10 Then
                         flg101 = True
                         If flg101 = True Then
                             Print #f1, "{RC012;" + PC1(I1) + "|}"
                         End If
                     End If

                     If I1 = 11 Then
                         flg111 = True
                         If flg111 = True Then
                             Print #f1, "{RC013;" + PC1(I1) + "|}"
                         End If
                     End If

                         SELPROCODE_RS.MoveNext
                         I1 = I1 + 1
                         GoTo code1
                End If
         Wend
         Print #f1, "{RB00;" + II + "|}"
         Print #f1, "{RC014;" + II + "|}"
         Print #f1, "{RC015;" + REC_DATE + "|}"
         Print #f1, "{RC016;" + REC_TIME + "|}"
         Print #f1, "{XS;I,0001,0002C4201|}"

End If

'Print #f1, "{RC003;1000|}"
'Print #f1, "{RC004;2000|}"
'Print #f1, "{RC005;3000|}"
'Print #f1, "{RC006;5000|}"
'Print #f1, "{RC007;6000|}"
'Print #f1, "{RC008;1001|}"
'Print #f1, "{RC009;2001|}"
'Print #f1, "{RC010;3001|}"
'Print #f1, "{RC011;5001|}"
'Print #f1, "{RC012;6001|}"
'Print #f1, "{RC013;4000|}"
'Print #f1, "{RB00;" + II + "|}"
'Print #f1, "{RB00;>6" + bar1 + ">5" + bar2 + ">6" + bar3 + "|}"
'Print #f1, "{RB00;ABCDEFGHIJKLM|}"
'Print #f1, "{RC014;" + II + "|}"
'Print #f1, "{RC014;ABCDEFGHIJKLM|}"
'Print #f1, "{RC015;" + REC_DATE + "|}"
'Print #f1, "{RC016;" + REC_TIME + "|}"
'Print #f1, "{XS;I,0001,0002C4201|}"

End Sub

