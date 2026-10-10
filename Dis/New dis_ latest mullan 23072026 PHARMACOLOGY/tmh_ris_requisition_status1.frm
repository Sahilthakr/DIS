VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form tmh_ris_requisition_status1 
   Caption         =   "REQUISITION STATUS"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport4 
      Left            =   1200
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport3 
      Left            =   720
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmd_emr 
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
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2400
      Width           =   735
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   0
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\DISREPORTSEC\disbm3aadd3.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport Cr_reqstat 
      Left            =   360
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CRYSTALREPORT1 
      Left            =   720
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1635
      Left            =   240
      TabIndex        =   26
      Top             =   4560
      Visible         =   0   'False
      Width           =   11775
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   11775
      Begin VB.Frame emppatient 
         Height          =   1095
         Left            =   1080
         TabIndex        =   30
         Top             =   480
         Width           =   1455
         Begin VB.OptionButton opt_patient 
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
            Height          =   255
            Left            =   240
            TabIndex        =   1
            Top             =   240
            Width           =   975
         End
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
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.OptionButton Opt_all 
         Caption         =   "&All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   600
         Width           =   645
      End
      Begin VB.OptionButton Opt_case 
         Caption         =   "Case &No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cbo_req 
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
         Height          =   315
         Left            =   9150
         TabIndex        =   17
         Top             =   1170
         Width           =   2295
      End
      Begin VB.ComboBox cbo_modality 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4590
         TabIndex        =   5
         Top             =   600
         Width           =   675
      End
      Begin VB.CommandButton CMD_VIEW2 
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
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4560
         TabIndex        =   18
         Top             =   240
         Width           =   6255
      End
      Begin MSMask.MaskEdBox MaskEd_to_date 
         Height          =   375
         Left            =   9120
         TabIndex        =   7
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
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
      Begin MSMask.MaskEdBox MaskEd_from_date 
         Height          =   375
         Left            =   6720
         TabIndex        =   6
         Top             =   600
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
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
      Begin MSMask.MaskEdBox msk_caseno 
         Height          =   375
         Left            =   4590
         TabIndex        =   20
         ToolTipText     =   "ENTER THE CASE NUMBER TO DISPLAY PATIENT DETAILS"
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
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
         Mask            =   ">??/#####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5520
         TabIndex        =   25
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8160
         TabIndex        =   24
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Requistion No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7500
         TabIndex        =   23
         Top             =   1200
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mod. Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3360
         TabIndex        =   22
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label LBL_NAME 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3840
         TabIndex        =   21
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3495
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   11775
      Begin MSFlexGridLib.MSFlexGrid msf_reqstat 
         Height          =   3255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "CLICK ON THE REQUISTION NO."
         Top             =   120
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   10
         BackColor       =   16777215
         ForeColor       =   0
         ForeColorFixed  =   192
         WordWrap        =   -1  'True
         Enabled         =   -1  'True
         AllowUserResizing=   1
         FormatString    =   $"tmh_ris_requisition_status1.frx":0000
      End
   End
   Begin VB.Frame Frame4 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   7080
      Width           =   11775
      Begin VB.CommandButton cmd_view 
         Caption         =   "&View"
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
         Left            =   2760
         MouseIcon       =   "tmh_ris_requisition_status1.frx":00F2
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         MouseIcon       =   "tmh_ris_requisition_status1.frx":03FC
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmd_exit 
         Caption         =   "&Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         MouseIcon       =   "tmh_ris_requisition_status1.frx":0706
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         MouseIcon       =   "tmh_ris_requisition_status1.frx":0A10
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.TextBox txt_time 
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   10590
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txt_date 
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   10590
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin Crystal.CrystalReport CrystalReport5 
      Left            =   0
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\DISREPORTSEC\disbm3aadd3.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport6 
      Left            =   360
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\DISREPORTSEC\disbm3aadd3.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      Caption         =   "05/11/21"
      Height          =   375
      Left            =   0
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   855
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
      Height          =   495
      Left            =   960
      TabIndex        =   31
      Top             =   120
      Width           =   11655
   End
   Begin VB.Label Label3 
      Caption         =   "REQUISITION STATUS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   4560
      TabIndex        =   29
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   9960
      TabIndex        =   28
      Top             =   1350
      Width           =   540
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   9960
      TabIndex        =   27
      Top             =   990
      Width           =   510
   End
End
Attribute VB_Name = "tmh_ris_requisition_status1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VAR_REQNNO As String
Dim var_caseno As String
Dim enadis As String
Dim cno_pad As String
Dim str As String
Dim flagall As String
Dim flagcase As String
Dim srno As Integer
Dim appt As String
Dim exam_flag As Boolean
Dim RS As rdoResultset
Dim rq_modcd As rdoQuery
Dim rs_modcd As rdoResultset

'ADDED BY MILIND ON 11-02-06
Dim qrfrep_qr As rdoQuery
Dim rsfrep_rs As rdoResultset
Dim adndm_qr As rdoQuery
Dim adndm_rs As rdoResultset
'upto here

Dim SELDADDENDUMFAQR As rdoQuery
Dim SELDADDENDUMFARS As rdoResultset

Dim SELDFN_QR As rdoQuery
Dim SELDFN_RS As rdoResultset

Dim mbqrysts_qr As rdoQuery
Dim mbqrysts_rs As rdoResultset

Dim rq_req As rdoQuery
Dim rs_req As rdoResultset

Dim rq_allreq As rdoQuery
Dim RQ_ALLREQ1 As rdoQuery
Dim rs_allreq As rdoResultset

Dim rq_grid_det As rdoQuery
Dim rs_grid_det As rdoResultset
Dim SELRISREPORT_RQ As rdoQuery
Dim SELRISREPORT_RS As rdoResultset

Dim SELRCLINIMPR_RQ As rdoQuery
Dim SELRCLINIMPR_RS As rdoResultset

Dim SELRADDENDUM_RQ As rdoQuery
Dim SELRADDENDUM_RS As rdoResultset
Dim W1 As Application
Dim D1 As Document
Dim ADDFLAG As Boolean
Dim REPFLAG As Boolean
Dim IMPFLAG As Boolean
Dim SELRMODALITY_QR As rdoQuery
Dim SELRMODALITY_RS As rdoResultset
Dim RIS1_QR As rdoQuery
Dim RIS1_RS As rdoResultset
Dim selrrequisitnqr As rdoQuery
Dim selrrequisitnrs As rdoResultset

Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset
Dim SELMA301_QR As rdoQuery
Dim SELMA301_RS As rdoResultset

Dim SELDFREPORT_QR As rdoQuery
Dim SELDFREPORT_RS As rdoResultset

Dim SELDADDENDUMQR As rdoQuery
Dim SELDADDENDUMRS As rdoResultset

Dim SELDREQNOQR1 As rdoQuery
Dim SELDREQNORS1 As rdoResultset
Dim SELSITEDETRQR1 As rdoQuery
Dim SELSITEDETRRS As rdoResultset

Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset
Dim SELTA101_FLAG As Boolean

Dim DMULPROCE_QR As rdoQuery
Dim DMULPROCE_RS As rdoResultset
Dim DMULPROCEFLAG  As Boolean
Dim IMFLAG As Boolean
Dim IMFLAG1 As Boolean

Dim SELSTREPORTQR As rdoQuery
Dim SELSTREPORTRS As rdoResultset

Dim SELECREPORTQR As rdoQuery
Dim SELECREPORTRS As rdoResultset
Dim SELECREPORTFLAG As Boolean

Dim SELDTMNO_QR As rdoQuery
Dim SELDTMNO_RS As rdoResultset
Dim SELDTMNO_FLAG As Boolean

Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITNFLAG1 As Boolean

Dim SELEMPBASINFO_QR As rdoQuery
Dim SELEMPBASINFO_RS As rdoResultset
Dim SELEMPBASINFOFLAG As Boolean

Dim SELDBMAREPORT_QR As rdoQuery
Dim SELDBMAREPORT_RS As rdoResultset
Dim SELDBMAREPORT_FLAG As Boolean

''FOR SP REPORT
Dim RQ_ALLSP1 As rdoQuery


Dim RQ_ALLSP As rdoQuery
Dim RS_ALLSP As rdoResultset
'Dim FLAG_ALLSP As Boolean

''DTM BLANK REPORT CHECK
Dim SELDTMTAB_QR As rdoQuery
Dim SELDTMTAB_RS As rdoResultset
Dim SELDTMTAB_FLAG As Boolean

''FOR DISLNKUSR
Dim LNK_QR As rdoQuery
Dim LNK_RS As rdoResultset
Dim LNK_FLAG As Boolean
''HLREPORT
Dim SELDCBCREPORT_QR As rdoQuery
Dim SELDCBCREPORT_RS As rdoResultset
Dim SELDCBCREPORT_FLAG As Boolean
''HLREPORT
Dim SELDHEMCYTO_QR As rdoQuery
Dim SELDHEMCYTO_RS As rdoResultset
Dim REFDR_QR As rdoQuery
Dim REFDR_RS As rdoResultset
Dim REFDRFLAG1 As Boolean
Dim SELDBMASUB_QR As rdoQuery
Dim SELDBMASSUB_RS As rdoResultset

'-------- 30/12/08 Shraddha ----------
Dim REFDR_RS_N As rdoResultset
Dim REFDR_QR_N As rdoQuery

'-------- 09/10/09 Shraddha ----------
Dim SEL_DHEMCPDET_QR As rdoQuery
Dim SEL_DHEMCPDET_RS As rdoResultset

Dim SELCPDETVIEW_QR As rdoQuery
Dim SELCPDETVIEW_RS As rdoResultset
'-------------------------------------

'***** Shraddha 25/02/10 *****
Dim SEL_DMSPONANTI_QR As rdoQuery
Dim SEL_DMSPONANTI_RS As rdoResultset

Dim var_gzzmod As String

Dim SEL_DHINTER_QR As rdoQuery
Dim SEL_DHINTER_RS As rdoResultset
'***** Shraddha 04/01/11 *****
Dim sel_dhemcbc_qr As rdoQuery
Dim sel_dhemcbc_rs As rdoResultset

Dim SELSTAFFINFO_QR As rdoQuery '09/03/11
Dim SELSTAFFINFO_RS As rdoResultset '09/03/11

Dim SEL_DHEMCBCDET_QR As rdoQuery 'Shraddha 15/04/11
Dim SEL_DHEMCBCDET_RS As rdoResultset 'Shraddha 15/04/11

Dim SEL_DCBCREP1_QR As rdoQuery
Dim SEL_DCBCREP1_RS As rdoResultset
'*****************************

Dim CLPTIME As String

''' INTERET EXPLOERER VARIBLES FOR PACS
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cbo_modality_Click()
If Trim(cbo_modality) = "" Then
Exit Sub
End If
MaskEd_from_date.Enabled = True
   MaskEd_from_date.SetFocus
   'txt_msg.Text = "ENTER FROM DATE AND TO DATE"
    MaskEd_to_date.Enabled = True
   MaskEd_to_date.Enabled = True
   MaskEd_from_date.Text = Format(DATE11, "dd/mm/yyyy")
   MaskEd_to_date.Text = Format(DATE11, "dd/mm/yyyy")
If cbo_modality.Text = "BC" Or cbo_modality.Text = "EL" Then
msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE            |^ CASE NO.    |^                  MEMO                 |^    PAYMENT   |^   COLLECTION    |^       RECEIVED        |^       REPORT            |^ REQNO.             |^   LAB NO.  "
End If

'***** Added by Shraddha Pandit on 25/02/10 *****
If cbo_modality.Text = "BA" Or cbo_modality.Text = "SE" Or cbo_modality.Text = "CM" Or cbo_modality.Text = "MB" Or cbo_modality.Text = "MO" Or cbo_modality.Text = "MC" Or cbo_modality.Text = "XX" Then
    msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE            |^ SPECIMEN  / CASE NO.    |^                  MEMO                 |^    PAYMENT   |^   COLLECTION    |^       RECEIVED        |^       REPORT            |^ PENDING REMARKS.             |^ PATH/LAB NO."
    'msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE            |^ CASE NO.    |^                  MEMO                 |^    PAYMENT   |^   COLLECTION    |^       RECEIVED        |^       REPORT            |^ PENDING REMARKS.             |^ PATH/LAB NO."
Else
    msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE            |^ CASE NO.    |^                  MEMO                 |^    PAYMENT   |^   COLLECTION    |^       RECEIVED        |^       REPORT            |^ PENDING REMARKS.             |^ PATH/LAB NO."
End If
'************************************************

End Sub

Private Sub cbo_req_Click()
If Trim(cbo_req) = "" Then
Exit Sub
End If
Dim flag As Boolean

If cbo_req.ListIndex = 0 Then
srno = 1
 msf_reqstat.Rows = 1
 End If
For i = 1 To msf_reqstat.Rows - 1
    msf_reqstat.row = i
    msf_reqstat.col = 1
    If msf_reqstat.Text = (cbo_req.Text) Then
        flag = True
    End If
Next i
If flag = False Then

rpt:

If cbo_req.ListIndex = 0 Then
    rq_grid_det(0) = req_chglen(cbo_req.List(srno))
Else
    rq_grid_det(0) = req_chglen(cbo_req.Text)
End If

Dim reqnov As String
reqnov = rq_grid_det(0)
            rs_grid_det.MoreResults
            rs_grid_det.Requery
            
            
             If rs_grid_det.RowCount <> 0 Then
            ' srno = 1
                While Not rs_grid_det.EOF
                  i = rs_grid_det.RowCount
                                    
                                    
                                    If Trim(rs_grid_det(1)) = "M" Then
                                        appt = "No"
                                        Memo = "Yes"
                                        payment = "No"
                                        EXAM = "No"
                                        REPORT = "No"
                                    End If
                                    
                                                                        
                                    If Trim(rs_grid_det(1)) = "P" Then
                                       'If Trim(rs_grid_det("RECPTNO")) = "NOCHARGE" Or Trim(rs_grid_det("RECPTNO")) = "TRUST" Or Trim(rs_grid_det("RECPTNO")) = "COMPANY" Then
                                        If IsNull(Trim(rs_grid_det("RECPTNO"))) Then
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "YES"
                                        EXAM = "No"
                                        REPORT = "No"
                                        Else
                                            If Trim(rs_grid_det("RECPTNO")) = "NOCHARGE" Or Trim(rs_grid_det("RECPTNO")) = "COMPANY" Or Trim(rs_grid_det("RECPTNO")) = "TRUST" Or Trim(rs_grid_det("RECPTNO")) = "IPATIENT" Or Trim(rs_grid_det("RECPTNO")) = "ITRUST" Or Trim(rs_grid_det("RECPTNO")) = "ICOMPANY" Or Trim(rs_grid_det("RECPTNO")) = "VPP" Or Trim(rs_grid_det("RECPTNO")) = "LR" Or Trim(rs_grid_det("RECPTNO")) = "BAR" Or Trim(rs_grid_det("RECPTNO")) = "DrOp" Then
                                            appt = "No"
                                            Memo = "Yes"
                                            payment = "Yes"
                                            EXAM = "No"
                                            REPORT = "No"
                                            Else
                                                If IsNull(Trim(rs_grid_det("EXAMDT"))) Then
                                                appt = "No"
                                                Memo = "Yes"
                                                payment = "Yes"
                                                EXAM = "No"
                                                REPORT = "No"
                                                Else
                                                appt = "Yes"
                                                Memo = "Yes"
                                                payment = "Yes"
                                                EXAM = "Yes"
                                                REPORT = "No"
                                                End If
                                            
                                            End If
                                        End If
                                   End If
                                   If Trim(rs_grid_det(1)) = "C" Then
                                       
                                        appt = "Yes"
                                        Memo = "YES"
                                        payment = "YES"
                                        EXAM = "No"
                                        REPORT = "No"
                                       
                                       
                                    End If
                                    
                                    If Trim(rs_grid_det(1)) = "E" Then
                                      If IsNull(Trim(rs_grid_det("RECPTNO"))) Then
                                        
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yes"
                                        Else
                                        If Trim(rs_grid_det("RECPTNO")) = "NOCHARGE" Or Trim(rs_grid_det("RECPTNO")) = "COMPANY" Or Trim(rs_grid_det("RECPTNO")) = "TRUST" Or Trim(rs_grid_det("RECPTNO")) = "IPATIENT" Or Trim(rs_grid_det("RECPTNO")) = "ITRUST" Or Trim(rs_grid_det("RECPTNO")) = "ICOMPANY" Or Trim(rs_grid_det("RECPTNO")) = "VPP" Or Trim(rs_grid_det("RECPTNO")) = "LR" Or Trim(rs_grid_det("RECPTNO")) = "BAR" Or Trim(rs_grid_det("RECPTNO")) = "DrOp" Then
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yes"
                                        Else
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "Yes"
                                        EXAM = "Yes"
                                        REPORT = "Yes"
                                        End If
                                        End If
                                   End If
                                    
                                    If Trim(rs_grid_det(1)) = "F" Then
                                       If IsNull(Trim(rs_grid_det("RECPTNO"))) Then
                                       
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yesf"
                                        '***** Shraddha Pandit 22/07/09 *****
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "SP" Then
                                                REPORT = "HPR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "SP" Then
                                                REPORT = "HPR"
                                            End If
                                        End If
                                        '************************************
                                        'Shraddha Pandit 04/08/09
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "IH" Then
                                                REPORT = "IHR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "IH" Then
                                                REPORT = "IHR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "FA" Then
                                                REPORT = "FAR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "FA" Then
                                                REPORT = "FAR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "CG" Then
                                                REPORT = "CGR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "CG" Then
                                                REPORT = "CGR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "CN" Then
                                                REPORT = "CNR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "CN" Then
                                                REPORT = "CNR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "MP" Then
                                                REPORT = "MPR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "MP" Then
                                                REPORT = "MPR"
                                            End If
                                        End If
                                        '************************************
                                        
                                       Else
                                        If Trim(rs_grid_det("RECPTNO")) = "NOCHARGE" Or Trim(rs_grid_det("RECPTNO")) = "COMPANY" Or Trim(rs_grid_det("RECPTNO")) = "TRUST" Or Trim(rs_grid_det("RECPTNO")) = "IPATIENT" Or Trim(rs_grid_det("RECPTNO")) = "ITRUST" Or Trim(rs_grid_det("RECPTNO")) = "ICOMPANY" Or Trim(rs_grid_det("RECPTNO")) = "VPP" Or Trim(rs_grid_det("RECPTNO")) = "LR" Or Trim(rs_grid_det("RECPTNO")) = "BAR" Or Trim(rs_grid_det("RECPTNO")) = "DrOp" Then
                                         appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yesf"
                                        '***** Shraddha Pandit 22/07/09 *****
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "SP" Then
                                                REPORT = "HPR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "SP" Then
                                                REPORT = "HPR"
                                            End If
                                        End If
                                        '************************************
                                        'Shraddha Pandit 04/08/09
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "IH" Then
                                                REPORT = "IHR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "IH" Then
                                                REPORT = "IHR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "FA" Then
                                                REPORT = "FAR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "FA" Then
                                                REPORT = "FAR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "CG" Then
                                                REPORT = "CGR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "CG" Then
                                                REPORT = "CGR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "CN" Then
                                                REPORT = "CNR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "CN" Then
                                                REPORT = "CNR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "MP" Then
                                                REPORT = "MPR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "MP" Then
                                                REPORT = "MPR"
                                            End If
                                        End If
                                        '************************************
                                        Else
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "YES"
                                        EXAM = "Yes"
                                        REPORT = "Yesf"
                                       '***** Shraddha Pandit 22/07/09 *****
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "SP" Then
                                                REPORT = "HPR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "SP" Then
                                                REPORT = "HPR"
                                            End If
                                        End If
                                        '***********************************
                                        'Shraddha Pandit 04/08/09
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "IH" Then
                                                REPORT = "IHR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "IH" Then
                                                REPORT = "IHR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "FA" Then
                                                REPORT = "FAR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "FA" Then
                                                REPORT = "FAR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "CG" Then
                                                REPORT = "CGR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "CG" Then
                                                REPORT = "CGR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "CN" Then
                                                REPORT = "CNR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "CN" Then
                                                REPORT = "CNR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "MP" Then
                                                REPORT = "MPR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "MP" Then
                                                REPORT = "MPR"
                                            End If
                                        End If
                                        '************************************
                                        End If
                                       End If
                                    End If
                                    If Trim(rs_grid_det(1)) = "T" Then
                                      If IsNull(Trim(rs_grid_det("RECPTNO"))) Then
                                        
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yest"
                                        '***** Shraddha Pandit 22/07/09 *****
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "SP" Then
                                                REPORT = "HPR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "SP" Then
                                                REPORT = "HPR"
                                            End If
                                        End If
                                        '***********************************
                                        'Shraddha Pandit 04/08/09
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "IH" Then
                                                REPORT = "IHR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "IH" Then
                                                REPORT = "IHR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "FA" Then
                                                REPORT = "FAR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "FA" Then
                                                REPORT = "FAR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "CG" Then
                                                REPORT = "CGR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "CG" Then
                                                REPORT = "CGR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "CN" Then
                                                REPORT = "CNR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "CN" Then
                                                REPORT = "CNR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "MP" Then
                                                REPORT = "MPR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "MP" Then
                                                REPORT = "MPR"
                                            End If
                                        End If
                                        '************************************
                                      Else
                                      If Trim(rs_grid_det("RECPTNO")) = "NOCHARGE" Or Trim(rs_grid_det("RECPTNO")) = "COMPANY" Or Trim(rs_grid_det("RECPTNO")) = "TRUST" Or Trim(rs_grid_det("RECPTNO")) = "IPATIENT" Or Trim(rs_grid_det("RECPTNO")) = "ITRUST" Or Trim(rs_grid_det("RECPTNO")) = "ICOMPANY" Or Trim(rs_grid_det("RECPTNO")) = "VPP" Or Trim(rs_grid_det("RECPTNO")) = "LR" Or Trim(rs_grid_det("RECPTNO")) = "BAR" Or Trim(rs_grid_det("RECPTNO")) = "DrOp" Then
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yest"
                                        '***** Shraddha Pandit 22/07/09 *****
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "SP" Then
                                                REPORT = "HPR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "SP" Then
                                                REPORT = "HPR"
                                            End If
                                        End If
                                        '************************************
                                        'Shraddha Pandit 04/08/09
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "IH" Then
                                                REPORT = "IHR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "IH" Then
                                                REPORT = "IHR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "FA" Then
                                                REPORT = "FAR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "FA" Then
                                                REPORT = "FAR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "CG" Then
                                                REPORT = "CGR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "CG" Then
                                                REPORT = "CGR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "CN" Then
                                                REPORT = "CNR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "CN" Then
                                                REPORT = "CNR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "MP" Then
                                                REPORT = "MPR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "MP" Then
                                                REPORT = "MPR"
                                            End If
                                        End If
                                        '************************************
                                        Else
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "YES"
                                        EXAM = "Yes"
                                        REPORT = "Yest"
                                        '***** Shraddha Pandit 22/07/09 *****
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "SP" Then
                                                REPORT = "HPR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "SP" Then
                                                REPORT = "HPR"
                                            End If
                                        End If
                                        '***********************************
                                        'Shraddha Pandit 04/08/09
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "IH" Then
                                                REPORT = "IHR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "IH" Then
                                                REPORT = "IHR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "FA" Then
                                                REPORT = "FAR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "FA" Then
                                                REPORT = "FAR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "CG" Then
                                                REPORT = "CGR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "CG" Then
                                                REPORT = "CGR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "CN" Then
                                                REPORT = "CNR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "CN" Then
                                                REPORT = "CNR"
                                            End If
                                        End If
                                        
                                        If Trim(cbo_req.Text) = "ALL" Then
                                            If Mid(Trim(cbo_req.List(srno)), 5, 2) = "MP" Then
                                                REPORT = "MPR"
                                            End If
                                        Else
                                            If Mid(Trim(cbo_req.Text), 5, 2) = "MP" Then
                                                REPORT = "MPR"
                                            End If
                                        End If
                                        '************************************
                                      End If
                                    End If
                                    End If
'                                    If Trim(rs_grid_det(1)) = "R" Then
'                                        appt = "Yes"
'                                        Memo = "Yes"
'                                        payment = "Yes"
'                                        EXAM = "Yes"
'                                        REPORT = "Yes"
'                                    End If
                                    If IsNull(rs_grid_det(1)) Then
                                        appt = "No"
                                        Memo = "No"
                                        payment = "No"
                                        EXAM = "No"
                                        REPORT = "No"
                                   End If
            
            
            '         msf_reqstat.AddItem srno & vbTab & cbo_req.Text & vbTab & rs_grid_det(1) & vbTab & Format(rs_grid_det(2), "dd/mm/yyyy") & vbTab & rs_grid_det(3) & vbTab & rs_grid_det(4)
                        'msf_reqstat.AddItem srno & vbTab & req_chglen(cbo_req.Text) & vbTab & rs_grid_det(1) & vbTab & appt & vbTab & Memo & vbTab & payment & vbTab & exam & vbTab & REPORT
                        
                        
                        msf_reqstat.Rows = msf_reqstat.Rows + 1
                             msf_reqstat.row = srno
                             msf_reqstat.col = 0
                             msf_reqstat.TextMatrix(srno, 0) = srno
                             If cbo_req.ListIndex = 0 Then
                             msf_reqstat.TextMatrix(srno, 1) = Trim(cbo_req.List(srno)) & " " & Trim(rs_grid_det("reqndt"))
                             Else
                             msf_reqstat.TextMatrix(srno, 1) = Trim(cbo_req) & " " & Trim(rs_grid_det("reqndt"))
                             End If
                             'msf_reqstat.TextMatrix(srno, 2) = msk_caseno.FormattedText
                             
                             '***** Shraddha 25/02/10 *****
                             If cbo_req.ListIndex = 0 Then
                                var_gzzmod = Mid(Trim(req_chglen(cbo_req.List(srno))), 4, 2)
                                SEL_DMSPONANTI_QR(0) = Trim(req_chglen(cbo_req.List(srno)))
                                Set SEL_DMSPONANTI_RS = SEL_DMSPONANTI_QR.OpenResultset(1, 3)
                             Else
                                var_gzzmod = Mid(Trim(req_chglen(cbo_req.Text)), 4, 2) ' Trim(req_chglen(cbo_req.Text))
                                SEL_DMSPONANTI_QR(0) = Trim(req_chglen(cbo_req.Text))
                                Set SEL_DMSPONANTI_RS = SEL_DMSPONANTI_QR.OpenResultset(1, 3)
                             End If
                                 If Trim(var_gzzmod) = "BA" Or Trim(var_gzzmod) = "SE" Or Trim(var_gzzmod) = "CM" Or Trim(var_gzzmod) = "MB" Or Trim(var_gzzmod) = "MO" Or Trim(var_gzzmod) = "MC" Or Trim(var_gzzmod) = "XX" Then
                                    'msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE            |^ SPECIMEN / CASENO    |^                  MEMO                 |^    PAYMENT   |^   COLLECTION    |^       RECEIVED        |^       REPORT            |^ PENDING REMARKS.             |^ PATH/LAB NO."
                                    'msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE            |^ SPECIMEN     |^                  MEMO                 |^    PAYMENT   |^   COLLECTION    |^       RECEIVED        |^       REPORT            |^ PENDING REMARKS.             |^ PATH/LAB NO."
                                    
    
                                    If Not SEL_DMSPONANTI_RS.EOF Then
                                        If IsNull(SEL_DMSPONANTI_RS("SPECIMEN")) = False Then
                                            msf_reqstat.ColWidth(2) = 3000
                                            'commented and added by shweta 04-07-2014
                                           ' msf_reqstat.TextMatrix(srno, 2) = Trim(SEL_DMSPONANTI_RS("SPECIMEN")) & "      |  " & msk_caseno.FormattedText
                                            msf_reqstat.TextMatrix(srno, 2) = Msk_caseno.FormattedText
                                            'msf_reqstat.TextMatrix(srno, 2) = Trim(SEL_DMSPONANTI_RS("SPECIMEN"))
                                        Else
                                            msf_reqstat.TextMatrix(srno, 2) = Msk_caseno.FormattedText
                                        End If
                                    End If
                                 Else
                                    msf_reqstat.TextMatrix(srno, 2) = Msk_caseno.FormattedText
                                 End If
                             
                             '*****************************
                             msf_reqstat.col = 3
                             If UCase(Memo) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                            msf_reqstat.TextMatrix(srno, 3) = Trim(rs_grid_det("MEMONO")) & " " & Trim(rs_grid_det("MEMODT")) & ""
                             End If
                            msf_reqstat.col = 4
                            
                           
                        
                            msf_reqstat.ColWidth(4) = 1900
                            If UCase(payment) = "YES" Then
                                msf_reqstat.CellBackColor = vbGreen
                                msf_reqstat.TextMatrix(srno, 4) = rs_grid_det("RECPTDT") & ""
                               
                                '***** Shraddha 17/12/09 *****
                                If Trim(rs_grid_det("REQNSTAT")) = "Z" Then
                                    msf_reqstat.TextMatrix(srno, 4) = rs_grid_det("RECPTDT") & " (REFUNDED)"
                                ElseIf Trim(rs_grid_det("REQNSTAT")) = "X" Then
                                    msf_reqstat.TextMatrix(srno, 4) = rs_grid_det("RECPTDT") '& "(X)"
                                ElseIf Trim(rs_grid_det("REQNSTAT")) = "U" Then
                                    msf_reqstat.TextMatrix(srno, 4) = rs_grid_det("RECPTDT") & " (Transferred)"
                                End If
                                '*****************************
                            Else
                               msf_reqstat.CellBackColor = vbGreen
                               If Trim(rs_grid_det("Recptno")) = "NOCHARGE" Then
                                msf_reqstat.TextMatrix(srno, 4) = "NOCHARGE"
                                
                                '***** Shraddha 05/01/10 *****
                                If Trim(rs_grid_det("REQNSTAT")) = "Z" Then
                                   msf_reqstat.TextMatrix(srno, 4) = "NOCHARGE" & " (REFUNDED)"
                                End If
                                
                               End If

                               If Trim(rs_grid_det("RECPTNO")) = "TRUST" Then
                                msf_reqstat.TextMatrix(srno, 4) = "TRUST"
                                '***** Shraddha 05/01/10 *****
                                If Trim(rs_grid_det("REQNSTAT")) = "Z" Then
                                   msf_reqstat.TextMatrix(srno, 4) = "TRUST" & " (REFUNDED)"
                                End If
                                
                               End If
                               
                               If Trim(rs_grid_det("RECPTNO")) = "COMPANY" Then
                                msf_reqstat.TextMatrix(srno, 4) = "COMPANY"
                                '***** Shraddha 05/01/10 *****
                                If Trim(rs_grid_det("REQNSTAT")) = "Z" Then
                                   msf_reqstat.TextMatrix(srno, 4) = "COMPANY" & " (REFUNDED)"
                                End If
                                
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "IPATIENT" Then
                                msf_reqstat.TextMatrix(srno, 4) = "IPATIENT"
                                '***** Shraddha 05/01/10 *****
                                If Trim(rs_grid_det("REQNSTAT")) = "Z" Then
                                   msf_reqstat.TextMatrix(srno, 4) = "IPATIENT" & " (REFUNDED)"
                                End If
                                
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "ITRUST" Then
                                msf_reqstat.TextMatrix(srno, 4) = "ITRUST"
                                '***** Shraddha 05/01/10 *****
                                If Trim(rs_grid_det("REQNSTAT")) = "Z" Then
                                   msf_reqstat.TextMatrix(srno, 4) = "ITRUST" & " (REFUNDED)"
                                End If
                                
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "ICOMPNY" Then
                                msf_reqstat.TextMatrix(srno, 4) = "ICOMPANY"
                                '***** Shraddha 05/01/10 *****
                                If Trim(rs_grid_det("REQNSTAT")) = "Z" Then
                                   msf_reqstat.TextMatrix(srno, 4) = "ICOMPANY" & " (REFUNDED)"
                                End If
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "VPP" Then
                                msf_reqstat.TextMatrix(srno, 4) = "VPP"
                                '***** Shraddha 05/01/10 *****
                                If Trim(rs_grid_det("REQNSTAT")) = "Z" Then
                                   msf_reqstat.TextMatrix(srno, 4) = "VPP" & " (REFUNDED)"
                                End If
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "LR" Then
                                msf_reqstat.TextMatrix(srno, 4) = "LR"
                                '***** Shraddha 05/01/10 *****
                                If Trim(rs_grid_det("REQNSTAT")) = "Z" Then
                                   msf_reqstat.TextMatrix(srno, 4) = "LR" & " (REFUNDED)"
                                End If
                                
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "BAR" Then
                                msf_reqstat.TextMatrix(srno, 4) = "BAR"
                                '***** Shraddha 05/01/10 *****
                                If Trim(rs_grid_det("REQNSTAT")) = "Z" Then
                                   msf_reqstat.TextMatrix(srno, 4) = "BAR" & " (REFUNDED)"
                                End If
                                
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "DrOp" Then
                                msf_reqstat.TextMatrix(srno, 4) = "DrOp"
                                '***** Shraddha 05/01/10 *****
                                If Trim(rs_grid_det("REQNSTAT")) = "Z" Then
                                   msf_reqstat.TextMatrix(srno, 4) = "DrOp" & " (REFUNDED)"
                                End If
                                
                               End If
                                
                               If IsNull(Trim(rs_grid_det("RECPTNO"))) Then
                                msf_reqstat.CellBackColor = vbWhite
                                msf_reqstat.TextMatrix(srno, 4) = " "
                                '***** Shraddha 05/01/10 *****
                                If Trim(rs_grid_det("REQNSTAT")) = "Z" Then
                                   msf_reqstat.TextMatrix(srno, 4) = " " & " (REFUNDED)"
                                End If
                                
                               End If
                                                    
                               


                             End If
                             msf_reqstat.col = 5
                             If UCase(appt) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                             msf_reqstat.TextMatrix(srno, 5) = rs_grid_det("APNTDT") & ""
                             End If
                             
                             msf_reqstat.col = 6
                             If Left(UCase(reqnov), 3) <> "LZZ" Then
                                If IsNull(Trim(rs_grid_det("SAMRECDATE"))) Then
                                        msf_reqstat.CellBackColor = vbWhite
                                        msf_reqstat.TextMatrix(srno, 6) = ""
                                    Else
                                      msf_reqstat.CellBackColor = vbGreen
                                       msf_reqstat.TextMatrix(srno, 6) = rs_grid_det("SAMRECDATE") & ""
                                    End If
                             Else
                                If UCase(EXAM) = "YES" Then
                                    msf_reqstat.CellBackColor = vbGreen
                                    If IsNull(Trim(rs_grid_det("EXAMDT"))) Then
                                        msf_reqstat.CellBackColor = vbWhite
                                        msf_reqstat.TextMatrix(srno, 6) = ""
                                    Else
                                       msf_reqstat.TextMatrix(srno, 6) = rs_grid_det("EXAMDT") & ""
                                    End If
                                End If
                             End If
                            msf_reqstat.col = 7
                             If UCase(REPORT) = "YESF" Then
                                If Left(UCase(reqnov), 3) = "GZZ" Then 'CONDITION ADDED BY SHWETA 10-06-2013 AS PER MAIL BY DR.KELKAR MADAM (DATED ON 05-JUN-2013 AS MICROBIOLOGY REPORT)
                                  msf_reqstat.TextMatrix(srno, 7) = "RPT " & rs_grid_det("FREPORTDT") & ""
                                Else
                                 msf_reqstat.TextMatrix(srno, 7) = "FNL " & rs_grid_det("FREPORTDT") & ""
                                End If
                             msf_reqstat.CellBackColor = vbGreen
                             End If
                            
                              If UCase(REPORT) = "YEST" Then
                                If Left(UCase(reqnov), 3) = "GZZ" Then 'CONDITION ADDED BY SHWETA 10-06-2013 AS PER MAIL BY DR.KELKAR MADAM (DATED ON 05-JUN-2013 AS MICROBIOLOGY REPORT)
                                 msf_reqstat.TextMatrix(srno, 7) = "RPT " & rs_grid_det("PREPORTDT") & ""
                                Else
                                msf_reqstat.TextMatrix(srno, 7) = "PRO " & rs_grid_det("PREPORTDT") & ""
                                End If
                                msf_reqstat.CellBackColor = vbGreen
                             End If
                            '***** Shraddha Pandit 22/07/09 *****
                            If UCase(REPORT) = "HPR" Then
                                msf_reqstat.TextMatrix(srno, 7) = "HPR " & rs_grid_det("PREPORTDT") & ""
                                msf_reqstat.CellBackColor = vbGreen
                            End If
                            '************************************
                            'Shraddha Pandit 04/08/09
                            If UCase(REPORT) = "IHR" Then
                                msf_reqstat.TextMatrix(srno, 7) = "IHR " & rs_grid_det("PREPORTDT") & ""
                                msf_reqstat.CellBackColor = vbGreen
                            End If
                            
                            
                            If UCase(REPORT) = "FAR" Then
                                msf_reqstat.TextMatrix(srno, 7) = "FAR " & rs_grid_det("PREPORTDT") & ""
                                msf_reqstat.CellBackColor = vbGreen
                            End If
                            
                            If UCase(REPORT) = "CGR" Then
                                msf_reqstat.TextMatrix(srno, 7) = "CGR " & rs_grid_det("PREPORTDT") & ""
                                msf_reqstat.CellBackColor = vbGreen
                            End If
                            
                            If UCase(REPORT) = "CNR" Then
                                msf_reqstat.TextMatrix(srno, 7) = "CNR " & rs_grid_det("PREPORTDT") & ""
                                msf_reqstat.CellBackColor = vbGreen
                            End If
                            
                            If UCase(REPORT) = "MPR" Then
                                msf_reqstat.TextMatrix(srno, 7) = "MPR " & rs_grid_det("PREPORTDT") & ""
                                msf_reqstat.CellBackColor = vbGreen
                            End If
                            '************************************
                        
                            msf_reqstat.col = 8
                            If Mid(rs_grid_det("REQNNO"), 4, 2) = "IH" Then
                                seldrequisitn_qr(0) = rs_grid_det("REQNNO")
                                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                                If seldrequisitn_rs.RowCount > 0 Then
                                    If Len(Trim(seldrequisitn_rs("SPIHREQNO"))) > 0 Then
                                        msf_reqstat.CellBackColor = vbGreen
                                        msf_reqstat.TextMatrix(srno, 8) = req_chglen(seldrequisitn_rs("spihreqno"))
                                        
                                    End If
                                End If
                            End If
                            If Mid(rs_grid_det("reqnno"), 4, 2) = "SP" Then
                                'SELDREQNOQR1(0) = rs_grid_det("REQNNO")
                                'Set SELDREQNORS1 = SELDREQNOQR1.OpenResultset(1, 3)
                                'If SELDREQNORS1.RowCount > 0 Then
                                 '   If Len(Trim(SELDREQNORS1("REQNO"))) > 0 Then
                                 '       msf_reqstat.CellBackColor = vbGreen
                                 '       msf_reqstat.TextMatrix(srno, 8) = req_chglen(SELDREQNORS1("REQNO"))
                                        
                                        
                                 '   End If
                            '''
                             ' THE CODE SLICE IS COMMENTED BECAUSE THE REQNO IS NOT REQUIRED BY PENDING REMARKS
      ' FIELD FOR THE SP IS REQUIRED FROM SITEDETAIL TABLE OF DISLIB
                                SELSITEDETRQR1(0) = rs_grid_det("reqnno")
                           
                                Set SELSITEDETRRS = SELSITEDETRQR1.OpenResultset(1, 3)
                                If SELSITEDETRRS.RowCount > 0 Then
                                    msf_reqstat.CellBackColor = vbGreen
                                    msf_reqstat.TextMatrix(srno, 8) = SELSITEDETRRS("FLD1")
                                End If
                            ''''
                               ' End If

                               
                               
                               
                               
                            End If
                            
'' PUTTING PICTURE LABEL FOR LZZ AND EXAM DONE FOR SUCH REQ
                            msf_reqstat.TextMatrix(srno, 8) = ""
                            If Mid(rs_grid_det("REQNNO"), 1, 3) = "LZZ" Then
                               If Len(Trim((msf_reqstat.TextMatrix(srno, 6)))) > 0 Then  '' exam date is being cheked
                                msf_reqstat.TextMatrix(srno, 8) = "PICTURE"
                                msf_reqstat.col = 1
                                msf_reqstat.CellBackColor = vbYellow
                               Else
                                 msf_reqstat.TextMatrix(srno, 8) = ""
                                 msf_reqstat.col = 1
                                 msf_reqstat.CellBackColor = vbWhite
                               End If
                                
                             End If
                             
                                 
                                 




''' PICTURE ENDS
                            
                            msf_reqstat.col = 9
                            seldrequisitn_qr(0) = rs_grid_det("REQNNO")
                            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                            If seldrequisitn_rs.RowCount > 0 Then
                                msf_reqstat.TextMatrix(srno, 9) = Trim(seldrequisitn_rs("PATHNO"))
                            End If
                        rs_grid_det.MoveNext
                        srno = srno + 1
                Wend
                
             End If
             
             If cbo_req.ListIndex = 0 Then
             If srno < cbo_req.ListCount Then
             GoTo rpt:
             End If
             
             End If
             
    Else
    MsgBox cbo_req.Text & "- REQUISION NUMBER IS ALREADY DISPLAYED."


    
End If
End Sub

Private Sub CMD_EMR_Click()
'If Len(Trim(Msk_caseno.Text)) > 0 Then
'    Dim VAR1
'    'VAR = Shell("C:\WINDOWS\Desktop\EMR 04.08.05.EXE" & Trim(Msk_caseno.FormattedText), vbNormalFocus)
'    VAR1 = Msk_caseno.FormattedText
'    VAR = Shell("C:\EMR1.EXE " & VAR1, vbNormalFocus)
'    AppActivate VAR
'Else
'    MsgBox "Please Enter the Case No."
'End If
If Len(Trim(Msk_caseno.Text)) > 0 Then
    VAR1 = Msk_caseno.FormattedText
    Dim strLink As String
    'strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
    strLink = "http://10.80.1.9/emrtmhact/TOPBAR.ASP?CASE_NO=" & VAR1 & ""
    Shell "c:\program files\internet explorer\iexplore " & strLink
Else
    MsgBox "Please Enter The Case No."
End If
End Sub

Private Sub cmd_exit_Click()
 button_ena_dis ("cancel")
  clear
Me.Hide
DIS_MAIN.show
End Sub
Private Sub cmd_print_Click()
  button_ena_dis ("print")
  Dim str1 As String
  Dim str2 As String
  Dim STR3 As String
  
  On Error GoTo prtmsg:
  'commented by amarjeet on 05-11-2002 because of object not found error
  'txt_msg.Text = "PRINTING..."
DoEvents
Dim STR5 As String

msf_reqstat.row = 1
msf_reqstat.col = 1
'Dim VAR_REQNNO As String
Dim var_yr_f As Integer, var_yr_t As Integer
Dim var_mon_f As Integer, var_mon_t As Integer
Dim var_day_f As Integer, var_day_t As Integer
Dim var_date_f As String, var_date_t As String



VAR_REQNNO = req_chglen(msf_reqstat.Text)

'cr_reqstat.Connect = "UID=PATOOLS;PWD=PATOOLS;Database = EC_TMH"

If Opt_case.Value = True Then
    CrystalReport3.SelectionFormula = "{RREQUISITN.CASENO} = '" & Trim(cno_pad) & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
    CrystalReport3.ReportFileName = "C:\DISREPORTSEC\reqsts.rpt"
    
     str1 = ip
     str2 = DATE11
     STR3 = time1
    str1 = str1 & Space(5) & str2 & Space(5) & STR3
    str1 = Space(3) & str1
    CrystalReport3.Formulas(0) = "us_id = '" & str1 & "'"
Else
    var_date_f = Year(MaskEd_from_date) & "," & month(MaskEd_from_date) & "," & Day(MaskEd_from_date)
    var_date_t = Year(MaskEd_to_date) & "," & month(MaskEd_to_date) & "," & Day(MaskEd_to_date)
''
''   'cr_reqstat.SelectionFormula = "{RREQUISITN.REQNDT} >= Date (2001,09 , 25) AND {RREQUISITN.REQNDT} <= Date (2001,12 ,30 )"
''    cr_reqstat.SelectionFormula = "{RREQUISITN.MODALITYCD}= '" & Trim(cbo_modality) & "'"
''    MsgBox cr_reqstat.SelectionFormula
''    'cr_reqstat.SelectionFormula = "{RREQUISITN.REQNDT} >= Date (2001,09 , 25) AND {RREQUISITN.REQNDT} <= Date (2001,12 ,30 )"
''    cr_reqstat.Formulas(0) = "curtime = '" & txt_time.Text & "'"
''    cr_reqstat.Formulas(1) = "formdate = '" & MaskEd_from_date.FormattedText & "'"
''    cr_reqstat.Formulas(2) = "todate = '" & MaskEd_to_date.FormattedText & "'"
''    cr_reqstat.ReportFileName = App.Path & "\crreqsts.rpt"
     CrystalReport3.ReportFileName = "C:\DISREPORTSEC\reqsts1.rpt"
        str1 = ip
        str2 = DATE11
        STR3 = time1
        str1 = str1 & Space(5) & str2 & Space(5) & STR3
        STR5 = "F"
  'CR3.SelectionFormula = " {TA102.case_no} = '" & MCNO & "' and {TA102.APPT_DATE} = date( " & adt & ") and {ta102.inv_type} = '" & INVTP & "'"   ''modified 23899
        
      
       ' cr_reqstat.SelectionFormula = "{RREQUISITN.REQNDT} >= Date (2001,01 , 01) AND {RREQUISITN.REQNDT} <= Date (2002,01 ,01 )"
         CrystalReport3.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & Trim(cbo_modality) & " ' AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
        'cr_reqstat.SelectionFormula = "{RREQUISITN.MODALITYCD}='" & Trim(cbo_modality) & "'"
        'cr_reqstat.SelectionFormula = "{RREQUISITN.REQNSTAT} <> '" & "F" & "'"
        
        CrystalReport3.Formulas(3) = "us_id = '" & str1 & "'"
        CrystalReport3.Formulas(0) = "fromdt = '" & MaskEd_from_date.FormattedText & "'"
        CrystalReport3.Formulas(2) = "todt = '" & MaskEd_to_date.FormattedText & "'"
       
End If


CrystalReport3.WindowState = crptMaximized
CrystalReport3.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'cr_reqstat.Destination =crptToPrinter
CrystalReport3.Action = 1


'Cr_reqstat.Action = 1
'commented by amarjeet on 05-11-2002 because of object not found error
'txt_msg.Text = ""
Exit Sub

prtmsg: MsgBox Error(Err)
'commented by amarjeet on 05-11-2002 because of object not found error
'txt_msg.Text = ""

  
  
  
End Sub
Private Sub cmd_cancel_Click()
  button_ena_dis ("cancel")
  clear
  txt_name.Visible = False
  lbl_name.Visible = False ' to clear all the text fields on the form
  
'  msf_reqstat.Rows = 1
'txt_msg.Text = "CLICK ANY BUTTON TO PERFORM AN OPERATION"
  
End Sub
Public Function button_ena_dis(enadis)

    If enadis = "view" Then
       'CMDPRINT
        cmd_view.Enabled = False
        cmd_exit.Enabled = False
        cmd_cancel.Enabled = True
        cmd_print.Enabled = True
        
    End If
    
    If enadis = "print" Then
           'cmdVIEW

        cmd_exit.Enabled = True
        cmd_cancel.Enabled = True
        cmd_print.Enabled = False
        cmd_view.Enabled = True
    End If
       
    If enadis = "cancel" Then
        'cmdVIEW
        cmd_view.Enabled = True
        cmd_exit.Enabled = True
        cmd_cancel.Enabled = False
        cmd_print.Enabled = False
    End If
    
    If enadis = "exit" Then
    End If
    
  
   End Function


Private Sub cmd_view_Click()
clear
button_ena_dis ("view")
'txt_msg.Text = "SELECT ANY ALL OR CASE OPTION BUTTON"
Frame2.Enabled = True
OPT_ALL.Enabled = True
Opt_case.Enabled = True
OPT_ALL.Value = False
Opt_case.Value = False
Msk_caseno.Enabled = False
cbo_req.Enabled = False
MaskEd_from_date.Enabled = False
MaskEd_to_date.Enabled = False
End Sub


Private Sub Form_Activate()
'cmd_view.SetFocus
End Sub

Private Sub Form_Load()
Dim str As String
'Connection
LNK_FLAG = False
txt_name.Visible = False
lbl_name.Visible = False
srno = 1
Set RS = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107") '***2/2/99
DATE11 = RS("dt1") '**add 24/11/2k
time1 = Format(RS("TM1"), "hh:mm:ss") '***2/2/99
txt_date = Format(DATE11, "dd/mm/yyyy")
txt_time = time1
'   MaskEd_from_date.Text = Format(DATE11, "dd/mm/yyyy")
'   MaskEd_to_date.Text = Format(DATE11, "dd/mm/yyyy")
exam_flag = False

'***** Shraddha 25/02/10 *****
str = "SELECT * FROM DISLIB.DMSPONANTI WHERE REQNO=?"
Set SEL_DMSPONANTI_QR = rdb.CreateQuery("", str)
'*****************************

'***** Shraddha 10/12/10 *****
'str = "SELECT * FROM DISLIB.HINTERFACE WHERE REQNO=? and MACHINENO='A2120'"
str = "SELECT * FROM DISLIB.HINTERFACE WHERE REQNO=? and MACHINENO LIKE 'A212%'"
Set SEL_DHINTER_QR = rdb.CreateQuery("", str)
'*****************************

'***** SHRADDHA 15/04/11 *****
str = "SELECT * FROM DISLIB.DHEMCBCDET WHERE reqno=?"
Set SEL_DHEMCBCDET_QR = rdb.CreateQuery("", str)
'*****************************

'***** Shraddha 04/01/11 *****
str = "SELECT * FROM DISLIB.REPORTHLF1 WHERE reqno=? and OBSER_TYPE=?"
Set sel_dhemcbc_qr = rdb.CreateQuery("", str)

'***** Shraddha 09/03/11 *****
str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=?"
Set SELSTAFFINFO_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DCBCREP1 WHERE REQNO=?"
Set SEL_DCBCREP1_QR = rdb.CreateQuery("", str)
'*****************************

If exam_flag = False Then
   
    ''DTM BLANK REPORT CHEQUE
    str = "SELECT * FROM DISLIB.DTMTAB WHERE REQNO=?"
    Set SELDTMTAB_QR = rdb.CreateQuery("SELDTMTAB_STR11", str)
    
    ''DCBCREPORT FOR HL
    str = "SELECT * FROM DISLIB.DCBCREPORT WHERE REQNO=?"
    Set SELDCBCREPORT_QR = rdb.CreateQuery("SELDCBCREPORT_STR", str)
    
    ''DCBCREPORT FOR HL
    ''DISLNKUSR 'AM
    str = "SELECT DEPT_CODE FROM DISLIB.DISLNKUSR WHERE USR_ID=?"
    Set LNK_QR = rdb.CreateQuery("L", str)
    
    Frame4.Enabled = True
    'txt_msg.Enabled = False
    cmd_cancel.Enabled = False
    cmd_print.Enabled = False
    txt_date.Enabled = False
    txt_time.Enabled = False
    MaskEd_from_date.Enabled = False
    MaskEd_to_date.Enabled = False
    ' for selecting case no into combo box from RREQUISITN
    Set rq_modcd = rdb.CreateQuery("", "select distinct(modalitycd) from rislib.rmodality WHERE DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='FZZ'OR DEPT_CODE='FEC'OR DEPT_CODE='LZZ' OR DEPT_CODE='GZZ' OR DEPT_CODE='IZZ'OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' or DEPT_CODE = 'VZZ' ")
    Set rs_modcd = rq_modcd.OpenResultset(1, 1)
    cbo_modality.clear
    cbo_modality.AddItem "ALL"
    While rs_modcd.EOF = False
        cbo_modality.AddItem rs_modcd(0)
        rs_modcd.MoveNext
    Wend
 
    ' for selecting req no for the selected case no from RREQUISITN
    Set rq_req = rdb.CreateQuery("", "select REQNNO,reqndt from Rislib.RREQUISITN where caseno=? order by reqndt desc")
    rq_req(0) = "aaa"
    Set rs_req = rq_req.OpenResultset()
    
    ' for selecting  req no for option all between the dates
    ' Set rq_allreq = rdb.CreateQuery("allreq", "select REQNNO,REQNSTAT,APNTDT,MEMONO,RECPTNO,CASENO from RREQUISITN where REQNDT between cdate(?) and cdate(?)")'????
     Set rq_allreq = rdb.CreateQuery("", "select REQNNO,REQNSTAT,APNTDT,MEMONO,RECPTNO,CASENO,MEMODT,RECPTDT,EXAMDT,FREPORTDT,PREPORTDT,RECPTNO,reqndt from Rislib.RREQUISITN where (reqnstat is null or reqnstat<>'F') AND REQNDT between ? and  ? and modalitycd=? order by reqndt desc")
     Set RQ_ALLREQ1 = rdb.CreateQuery("", "select REQNNO,REQNSTAT,APNTDT,MEMONO,RECPTNO,CASENO,MEMODT,RECPTDT,EXAMDT,FREPORTDT,PREPORTDT,RECPTNO,reqndt from Rislib.RREQUISITN where (reqnstat is null or reqnstat<>'F') AND REQNDT between ? and  ? order by reqndt desc")
     rq_allreq(0) = "1999/01/01"  '"01/01/1999"
     rq_allreq(1) = "1999/01/10"
     rq_allreq(2) = "CT"
     Set rs_allreq = rq_allreq.OpenResultset(2, 3)
     
     ''FOR SP REPORT
     Set RQ_ALLSP = rdb.CreateQuery("allSP", "select REQNNO,REQNSTAT,APNTDT,MEMONO,RECPTNO,CASENO,MEMODT,RECPTDT,EXAMDT,FREPORTDT,PREPORTDT,RECPTNO,reqndt from Rislib.RREQUISITN where (reqnstat is null or reqnstat<>'T') AND REQNDT between ? and  ? and modalitycd=? order by reqndt desc")
     Set RQ_ALLSP1 = rdb.CreateQuery("allSP1", "select REQNNO,REQNSTAT,APNTDT,MEMONO,RECPTNO,CASENO,MEMODT,RECPTDT,EXAMDT,FREPORTDT,PREPORTDT,RECPTNO,reqndt from Rislib.RREQUISITN where (reqnstat is null or reqnstat<>'T') AND REQNDT between ? and  ? order by reqndt desc")
     RQ_ALLSP(0) = "1999/01/01"  '"01/01/1999"
     RQ_ALLSP(1) = "1999/01/10"
     RQ_ALLSP(2) = "CT"
     Set RS_ALLSP = RQ_ALLSP.OpenResultset(2, 3)
     
     'for displaying only micobiology requisition when login is made by microbiology department user
     'added by milind on 12-10-06
     str = "select reqnno,reqndt from rislib.rrequisitn where caseno=? and dept_code in(select dept_code from dislib.dislnkusr where usr_id=? and dept_code='GZZ') order by reqndt desc"
     Set mbqrysts_qr = rdb.CreateQuery("mbqrysts_qr", str)
     'upto here
     
     
    '  for selecting flexgrid details from RREQUISITN based on REQNSTAT details
     Set rq_grid_det = rdb.CreateQuery("grid_det", "select REQNNO,REQNSTAT,APNTDT,MEMONO,RECPTNO,MEMODT,EXAMDT,FREPORTDT,PREPORTDT ,FREPORTBY,RECPTDT,PATNSTAT,recptno,reqndt,ENTERBY,SAMRECDATE from Rislib.ZREQRTIME where REQNNO=?")
     rq_grid_det(0) = "aa"
     Set rs_grid_det = rq_grid_det.OpenResultset(2, 3)
     exam_flag = True
 
    ' STR = "SELECT * FROM RISLIB.RISREPORTS WHERE REQNNO=?"
    ' Set SELRISREPORT_RQ = rdb.CreateQuery("SELRISREPORT_RQ1", STR)
    ' SELRISREPORT_RQ(0) = "LZZCT01000000"
    ' Set SELRISREPORT_RS = SELRISREPORT_RQ.OpenResultset(2, 3)
     
     str = "SELECT * FROM RISLIB.RCLINIMPR WHERE REQNNO=?"
     Set SELRCLINIMPR_RQ = rdb.CreateQuery("SELRCLINIMPR_RQ1 ", str)
     SELRCLINIMPR_RQ(0) = "LZZCT01000000"
     Set SELRCLINIMPR_RS = SELRCLINIMPR_RQ.OpenResultset(2, 3)
     
     str = "SELECT * FROM DISLIB.DFNREPORT WHERE REQNO=?"
     Set SELDFN_QR = rdb.CreateQuery("SELDFN_QR", str)
                    
     str = "SELECT * FROM RISLIB.RADDENDUM WHERE REQNO=?"
     Set SELRADDENDUM_RQ = rdb.CreateQuery("SELRADDENDUM_RQ1", str)
     SELRADDENDUM_RQ(0) = "LZZCT01000000"
     Set SELRADDENDUM_RS = SELRADDENDUM_RQ.OpenResultset(2, 3)
     ADDFLAG = False
     REPFLAG = False
     IMPFLAG = False
     str = "SELECT * FROM RISLIB.RMODALITY WHERE MODALITYCD=?"
     Set SELRMODALITY_QR = rdb.CreateQuery("SELRMODALITY1_QR1", str)
     SELRMODALITY_QR(0) = "CT"
     Set SELRMODALITY_RS = SELRMODALITY_QR.OpenResultset(2, 3)
     
    ' STR = "SELECT * FROM RISLIB.RISAUTH WHERE USRGROUP=? AND SCRN_DESC=?"
    ' Set RIS1_QR = rdb.CreateQuery("RIS1_QR1", STR)
    ' RIS1_QR(0) = Trim(USRGRP)
    ' RIS1_QR(1) = "REQUISITION STATUS"
    ' Set RIS1_RS = RIS1_QR.OpenResultset(2, 3)
    ' If RIS1_RS.RowCount = 0 Then
    '     MsgBox "No Button Authorisation for this User Group & Screen"
    '     End
    ' End If
    
    ''STAFF DETAILS
    str = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
    Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO_QR111", str)
    
    
    ''REQNO SELECT FOR STAFF  '05-11-2012 SHWETA
    'str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=?"
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? ORDER BY REQNDT DESC"
    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR1BC", str)
    
    'Added by milind on 10-02-06 for displaying addendum and designation of the Dr.
    
    str = "select addend_no from dislib.daddendum where reqno=?"
    Set adndm_qr = rdb.CreateQuery("adndm_qr", str)
    
    str = "select freportby from Rislib.Rrequisitn where reqnNo=?"
    Set qrfrep_qr = rdb.CreateQuery("qrfrep_qr", str)
    
    'upto here
    
    ''DTMNO
    str = "SELECT * FROM DISLIB.DTMNO WHERE CASENO=?"
    Set SELDTMNO_QR = rdb.CreateQuery("SELDTMNO_STR111", str)
    
    str = "select * from Rislib.Rrequisitn where reqnNo=?"
    Set selrrequisitnqr = rdb.CreateQuery("selrrequisitnqr1", str)
    
    str = "select * from dislib.drequisitn where reqno=?"
    Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn1_qr1", str)
    
    str = "select * from patoolslib.ta1012 where case_no=?"
    Set selta1012_qr = rdb.CreateQuery("1selta10121_qr1", str)
    
    str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
    Set SELMA301_QR = rdb.CreateQuery("SELMA3011_QR1", str)
    
    str = "SELECT * FROM DISLIB.DFREPORT WHERE REQNO=?"
    Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT1_QR1", str)
    
    str = "SELECT * FROM DISLIB.DADDENDUM WHERE REQNO=?"
    Set SELDADDENDUMQR = rdb.CreateQuery("SELDADDENDUMQR1", str)
    
    str = "SELECT * FROM DISLIB.ADDENDUMFA WHERE REQNO=?"
    Set SELDADDENDUMFAQR = rdb.CreateQuery("SELDADDENDUMFAQR1", str)
    
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE SPIHREQNO=?"
        Set SELDREQNOQR1 = rdb.CreateQuery("S1ELDREQNOQR121", str)
        
    str = "SELECT * FROM DISLIB.SITEDETAIL WHERE REQNO=?"
       Set SELSITEDETRQR1 = rdb.CreateQuery("SELSITEDETRQR11", str)
       
    str = "SELECT * FROM PATOOLSLIB.TA101 WHERE CAsE_NO=?"
    Set SELTA101_QR = rdb.CreateQuery("RSELTA101_QR1", str)
    
    str = "SELECT pro_code FROM DISLIB.DMULPROCE WHERE REQNO=? "
    Set DMULPROCE_QR = rdb.CreateQuery("DMULPROCE_QR1TUMREPRQ", str)
    
    str = "SELECT * FROM DISLIB.DSTREPORT WHERE REQNO=?"
    Set SELSTREPORTQR = rdb.CreateQuery("SELSTREPORTQRSVZ", str)
    
    str = "SELECT * FROM DISLIB.DECREPORT WHERE REQNO=?"
    Set SELECREPORTQR = rdb.CreateQuery("SELDECREPORTQRXX1", str)
    
    '-------------- 30/12/08 Shraddha Pandit ----------
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR_N = rdb.CreateQuery("", str)
    '--------------------------------------------------
    
    QrySts = False 'Added by Shraddha Pandit on 07/09/09
    
    str = "SELECT * FROM DISLIB.DBMAREPORT WHERE REQNO=?"
    Set SELDBMAREPORT_QR = rdb.CreateQuery("SELDBMAREPORT_STR", str)
    str = "select * from dislib.dhemcyto1 where reqno = ?"
    Set SELDHEMCYTO_QR = rdb.CreateQuery("seledhemcyto11_qr", str)
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("REFDR_QR1HE1", str)
    str = "SELECT * FROM DISLIB.DBMASUBREP WHERE REQNO = ?"
    Set SELDBMASUB_QR = rdb.CreateQuery("DBMASUB11", str)
    
    '***** Shraddha 09/10/09 *****
    str = "SELECT * FROM DISLIB.DHEMCPDET WHERE REQNO = ?"
    Set SEL_DHEMCPDET_QR = rdb.CreateQuery("", str)
    
    str = "SELECT * FROM DISLIB.DHEMCPDET WHERE REQNO=?"
    Set SELCPDETVIEW_QR = rdb.CreateQuery("", str)
    '*****************************
    
     Frame2.Enabled = True
     OPT_ALL.Enabled = False
     Opt_case.Enabled = False
     button_ena_dis ("cancel")
     'CMD_VIEW.Enabled = True
     'cmd_view.SetFocus
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


Private Sub MaskEd_from_date_LostFocus()
OPT_ALL.Enabled = False
If InStr(MaskEd_from_date.Text, "_") > 0 Then
   MaskEd_from_date.Mask = ""
            MaskEd_from_date.Text = ""
            MaskEd_from_date.Mask = "##/##/####"
            Exit Sub
End If

If IsDate(GET_REVDATE(MaskEd_from_date.Text)) = False Then
             MsgBox "Please ENTER the correct DATE."
             MaskEd_from_date.Mask = ""
            MaskEd_from_date.Text = ""
            MaskEd_from_date.Mask = "##/##/####"
'             MaskEd_from_date.Text = ""
             MaskEd_from_date.SetFocus
             'Exit Sub
             Else
             MaskEd_to_date.Enabled = True
             MaskEd_to_date.SetFocus
    End If
End Sub
Public Sub CLPTM()
Dim RTLOG_QR As rdoQuery
Dim RTLOG_RS As rdoResultset

Dim VWDMULPROCE_QR As rdoQuery
Dim VWDMULPROCE_RS As rdoResultset


str = "SELECT SAMCOLDATE,SAMCOLTIME,ADPREPDATE,ADPREPTIME FROM RISLIB.RTIMELOG WHERE REQNNO=? "
Set RTLOG_QR = rdb.CreateQuery("RTM", str)

str = "SELECT * FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
Set VWDMULPROCE_QR = rdb.CreateQuery("VE", str)

VWDMULPROCE_QR(0) = req_chglen(Trim(cbo_req.Text))
Set VWDMULPROCE_RS = VWDMULPROCE_QR.OpenResultset(1, 3)

If Not VWDMULPROCE_RS.EOF Then
If VWDMULPROCE_RS("PRO_CODE") = "F833" Or VWDMULPROCE_RS("PRO_CODE") = "F834" Or VWDMULPROCE_RS("PRO_CODE") = "F835" Or VWDMULPROCE_RS("PRO_CODE") = "F836" Then

    RTLOG_QR(0) = req_chglen(Trim(cbo_req.Text))
    Set RTLOG_RS = RTLOG_QR.OpenResultset(1, 3)

    Dim t1, T2, D1, D2 As String
    Dim H1, H2, h, m1, m2, m As Integer

    D1 = Trim(RTLOG_RS("SAMCOLDATE"))
    t1 = Trim(RTLOG_RS("SAMCOLTIME"))
    D2 = Trim(RTLOG_RS("ADPREPDATE"))
    T2 = Trim(RTLOG_RS("ADPREPTIME"))

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
    


    If m2 < 60 And m2 > 0 Then
    m2 = 60 - m2
    H2 = H2 + 1
    End If
    m = m2 + m1
    If Right(t1, 2) = "PM" Then
        If H1 = 12 Then
         H1 = 0
        Else
         H1 = H1 + 12
        End If
    End If


    If Right(T2, 2) = "PM" Then
        If H2 = 12 Then
            H2 = 0
        Else
            H2 = H2 + 12
        End If
    End If

    If DateDiff("D", D2, D1) = 0 Then
        h = H1 - H2
    ElseIf DateDiff("D", D2, D1) = 1 Then
        h = 24 - H2 + H1
    ElseIf DateDiff("D", D2, D1) > 1 Then
        h = DateDiff("D", D2, D1) * 24
        h = h + (24 - H2 + H1)
    End If

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
End If
VWDMULPROCE_RS.Close
VWDMULPROCE_QR.Close

End Sub

Private Sub MaskEd_to_date_LostFocus()
Dim TEMP As String
'     MaskEd_to_date.Enabled = False
          
         
If InStr(MaskEd_to_date.Text, "_") > 0 Then
   MaskEd_to_date.Mask = ""
   MaskEd_to_date.Text = ""
   MaskEd_to_date.Mask = "##/##/####"
   Exit Sub
End If
If IsDate(GET_REVDATE(MaskEd_to_date.Text)) = False Then
    MaskEd_to_date.Mask = ""
    MaskEd_to_date.Text = ""
    MaskEd_to_date.Mask = "##/##/####"
    MsgBox "Please ENTER the  correct DATE."
    MaskEd_to_date.SetFocus
    Exit Sub
End If
If cbo_modality.Text = "SP" Then
    SPREPORT
    Exit Sub
End If
If cbo_modality.Text = "ALL" Then
     RQ_ALLREQ1(0) = GET_REVDATE(MaskEd_from_date.Text)
     RQ_ALLREQ1(1) = GET_REVDATE(MaskEd_to_date.Text)
     Set rs_allreq = RQ_ALLREQ1.OpenResultset(2, 3)
     cmd_print.Enabled = False
 Else
     rq_allreq(0) = GET_REVDATE(MaskEd_from_date.Text)
     rq_allreq(1) = GET_REVDATE(MaskEd_to_date.Text)
     rq_allreq(2) = Trim(cbo_modality)
     rs_allreq.MoreResults
     rs_allreq.Requery
     cmd_print.Enabled = False
End If
If rs_allreq.RowCount <> 0 Then
    cmd_print.Enabled = True
    rs_allreq.MoveFirst
    
    msf_reqstat.Rows = 1
    srno = 1
    While Not rs_allreq.EOF
'    While Not rs_grid_det.EOF
      i = rs_allreq.RowCount
                        
                        
                        If Trim(rs_allreq(1)) = "M" Then
                            appt = "No"
                            Memo = "Yes"
                            payment = "No"
                            EXAM = "No"
                            REPORT = "No"
                        End If
                        
                        If Trim(rs_allreq(1)) = "P" Then
                                       'If Trim(rs_grid_det("RECPTNO")) = "NOCHARGE" Or Trim(rs_grid_det("RECPTNO")) = "TRUST" Or Trim(rs_grid_det("RECPTNO")) = "COMPANY" Then
                                        If IsNull(Trim(rs_allreq("RECPTNO"))) Then
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "YES"
                                        EXAM = "No"
                                        REPORT = "No"
                                        Else
                                        If Trim(rs_allreq("RECPTNO")) = "NOCHARGE" Or Trim(rs_allreq("RECPTNO")) = "COMPANY" Or Trim(rs_allreq("RECPTNO")) = "TRUST" Or Trim(rs_allreq("RECPTNO")) = "ITRUST" Or Trim(rs_allreq("RECPTNO")) = "ICOMPANY" Or Trim(rs_allreq("RECPTNO")) = "IPATIENT" Or Trim(rs_allreq("RECPTNO")) = "VPP" Or Trim(rs_allreq("RECPTNO")) = "LR" Or Trim(rs_allreq("RECPTNO")) = "BAR" Or Trim(rs_allreq("RECPTNO")) = "DrOp" Then
                                        appt = "No"
                                        Memo = "Yes"
                                        payment = "Yes"
                                        EXAM = "No"
                                        REPORT = "No"
                                        Else
                                        If IsNull(Trim(rs_allreq("EXAMDT"))) Then
                                        appt = "No"
                                        Memo = "Yes"
                                        payment = "YES"
                                        EXAM = "No"
                                        REPORT = "No"
                                        Else
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "YES"
                                        EXAM = "Yes"
                                        REPORT = "No"
                                        End If
                                        End If
                                        End If
                                   End If
                                    If Trim(rs_allreq(1)) = "C" Then
                                        
                                        appt = "Yes"
                                        Memo = "YES"
                                        payment = "YES"
                                        EXAM = "No"
                                        REPORT = "No"
                                        
                                    End If
                                                
                                    If Trim(rs_allreq(1)) = "E" Then
                                      If IsNull(Trim(rs_allreq("RECPTNO"))) Then
                                        
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yes"
                                        Else
                                        If Trim(rs_allreq("RECPTNO")) = "NOCHARGE" Or Trim(rs_allreq("RECPTNO")) = "COMPANY" Or Trim(rs_allreq("RECPTNO")) = "TRUST" Or Trim(rs_allreq("RECPTNO")) = "ITRUST" Or Trim(rs_allreq("RECPTNO")) = "ICOMPANY" Or Trim(rs_allreq("RECPTNO")) = "IPATIENT" Or Trim(rs_allreq("RECPTNO")) = "VPP" Or Trim(rs_allreq("RECPTNO")) = "LR" Or Trim(rs_allreq("RECPTNO")) = "BAR" Or Trim(rs_allreq("RECPTNO")) = "DrOp" Then
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yes"
                                        Else
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "Yes"
                                        EXAM = "Yes"
                                        REPORT = "Yes"
                                        End If
                                        End If
                                   End If
                                    
                                    If Trim(rs_allreq(1)) = "F" Then
                                       If IsNull(Trim(rs_allreq("RECPTNO"))) Then
                                       
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yesf"
                                       Else
                                        If Trim(rs_allreq("RECPTNO")) = "NOCHARGE" Or Trim(rs_allreq("RECPTNO")) = "COMPANY" Or Trim(rs_allreq("RECPTNO")) = "TRUST" Or Trim(rs_allreq("RECPTNO")) = "ITRUST" Or Trim(rs_allreq("RECPTNO")) = "ICOMPANY" Or Trim(rs_allreq("RECPTNO")) = "IPATIENT" Or Trim(rs_allreq("RECPTNO")) = "VPP" Or Trim(rs_allreq("RECPTNO")) = "LR" Or Trim(rs_allreq("RECPTNO")) = "BAR" Or Trim(rs_allreq("RECPTNO")) = "DrOp" Then
                                         appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yesf"
                                        Else
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "YES"
                                        EXAM = "Yes"
                                        REPORT = "Yesf"
                                        End If
                                       End If
                                    End If
                                    If Trim(rs_allreq(1)) = "T" Then
                                      If IsNull(Trim(rs_allreq("RECPTNO"))) Then
                                        
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yest"
                                      Else
                                      If Trim(rs_allreq("RECPTNO")) = "NOCHARGE" Or Trim(rs_allreq("RECPTNO")) = "COMPANY" Or Trim(rs_allreq("RECPTNO")) = "TRUST" Or Trim(rs_allreq("RECPTNO")) = "ITRUST" Or Trim(rs_allreq("RECPTNO")) = "ICOMPANY" Or Trim(rs_allreq("RECPTNO")) = "IPATIENT" Or Trim(rs_allreq("RECPTNO")) = "VPP" Or Trim(rs_allreq("RECPTNO")) = "LR" Or Trim(rs_allreq("RECPTNO")) = "BAR" Or Trim(rs_allreq("RECPTNO")) = "DrOp" Then
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yest"
                                        Else
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "YES"
                                        EXAM = "Yes"
                                        REPORT = "Yest"
                                      End If
                                    End If
                                    End If
                        If IsNull(Trim(rs_allreq(1))) Then
                            appt = "No"
                            Memo = "No"
                            payment = "No"
                            EXAM = "No"
                            REPORT = "No"
                       End If


'         msf_reqstat.AddItem srno & vbTab & cbo_req.Text & vbTab & rs_grid_det(1) & vbTab & Format(rs_grid_det(2), "dd/mm/yyyy") & vbTab & rs_grid_det(3) & vbTab & rs_grid_det(4)
'            msf_reqstat.AddItem srno & vbTab & cbo_req.Text & vbTab & rs_grid_det(1) & vbTab & appt & vbTab & Memo & vbTab & payment & vbTab & exam & vbTab & report
'            rs_grid_det.MoveNext
'            srno = srno + 1
'                    If IsNull(Len(Trim(rs_allreq(2)))) Then
'                            appt = "no"
'                        Else
'                            appt = "yes"
'                        End If
                               'msf_reqstat.AddItem srno & vbTab & rs_allreq(0) & vbTab & rs_allreq(1) & vbTab & Format(rs_allreq(2), "dd/mm/yyyy") & vbTab & rs_allreq(3) & vbTab & rs_allreq(4)
'                               msf_reqstat.AddItem srno & vbTab & rs_allreq(0) & vbTab & rs_allreq(1) & vbTab & appt & vbTab & rs_allreq(3) & vbTab & rs_allreq(4)

                             'msf_reqstat.AddItem srno & vbTab & req_chglen(rs_allreq(0)) & vbTab & rs_allreq("CASENO") & vbTab & appt & vbTab & Memo & vbTab & payment & vbTab & exam & vbTab & Report
                             msf_reqstat.Rows = msf_reqstat.Rows + 1
                             msf_reqstat.row = srno
                             msf_reqstat.col = 0
                             msf_reqstat.TextMatrix(srno, 0) = srno
                             msf_reqstat.TextMatrix(srno, 1) = Trim(req_chglen(rs_allreq(0))) & "  " & Trim(rs_allreq("reqndt"))
                             'msf_reqstat.TextMatrix(srno, 2) = rs_allreq("caseno")
                             'Write code for Nmaterial for Microbiology 25/02/10
                             '***** Shraddha 25/02/10 *****
                             If cbo_modality.Text = "BA" Or cbo_modality.Text = "SE" Or cbo_modality.Text = "CM" Or cbo_modality.Text = "MB" Or cbo_modality.Text = "MO" Or cbo_modality.Text = "MC" Or cbo_modality.Text = "XX" Then
                                SEL_DMSPONANTI_QR(0) = Trim(rs_allreq("REQNNO"))
                                Set SEL_DMSPONANTI_RS = SEL_DMSPONANTI_QR.OpenResultset(1, 3)
                                
                                If Not SEL_DMSPONANTI_RS.EOF Then
                                    If IsNull(SEL_DMSPONANTI_RS("SPECIMEN")) = False Then
                                        msf_reqstat.ColWidth(2) = 3000
                                        msf_reqstat.TextMatrix(srno, 2) = Trim(SEL_DMSPONANTI_RS("SPECIMEN")) & "      |  " & rs_allreq("caseno")
                                    Else
                                        msf_reqstat.TextMatrix(srno, 2) = rs_allreq("caseno")
                                    End If
                                End If
                             Else
                                msf_reqstat.TextMatrix(srno, 2) = rs_allreq("caseno")
                             End If
                             '*****************************
                             msf_reqstat.col = 3
                             If UCase(Memo) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                             
                             msf_reqstat.TextMatrix(srno, 3) = rs_allreq("MEMONO") & " " & rs_allreq("MEMODT") & ""
                             End If
                            msf_reqstat.col = 4
                             If UCase(payment) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                             msf_reqstat.TextMatrix(srno, 4) = rs_allreq("RECPTDT") & ""
                             Else
                              msf_reqstat.CellBackColor = vbGreen
                               If Trim(rs_allreq("Recptno")) = "NOCHARGE" Then
                                msf_reqstat.TextMatrix(srno, 4) = "NOCHARGE"
                               End If

                               If Trim(rs_allreq("RECPTNO")) = "TRUST" Then
                                msf_reqstat.TextMatrix(srno, 4) = "TRUST"
                               End If
                               
                               If Trim(rs_allreq("RECPTNO")) = "COMPANY" Then
                                msf_reqstat.TextMatrix(srno, 4) = "COMPANY"
                               End If
                                
                               If Trim(rs_allreq("RECPTNO")) = "ICOMPANY" Then
                                msf_reqstat.TextMatrix(srno, 4) = "ICOMPANY"
                               End If
                               If Trim(rs_allreq("RECPTNO")) = "ITRUST" Then
                                msf_reqstat.TextMatrix(srno, 4) = "ITRUST"
                               End If
                               If Trim(rs_allreq("RECPTNO")) = "IPATIENT" Then
                                msf_reqstat.TextMatrix(srno, 4) = "IPATIENT"
                               End If
                               If Trim(rs_allreq("RECPTNO")) = "VPP" Then
                                msf_reqstat.TextMatrix(srno, 4) = "VPP"
                               End If
                               If Trim(rs_allreq("RECPTNO")) = "LR" Then
                                msf_reqstat.TextMatrix(srno, 4) = "LR"
                               End If
                               If Trim(rs_allreq("RECPTNO")) = "BAR" Then
                                msf_reqstat.TextMatrix(srno, 4) = "BAR"
                               End If
                               If Trim(rs_allreq("RECPTNO")) = "DrOp" Then
                                msf_reqstat.TextMatrix(srno, 4) = "DrOp"
                               End If
                               
                               If IsNull(Trim(rs_allreq("RECPTNO"))) Then
                                msf_reqstat.CellBackColor = vbWhite
                                msf_reqstat.TextMatrix(srno, 4) = " "
                               End If
                             
                                                      
                             
                             End If
                             msf_reqstat.col = 5
                             If UCase(appt) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                             msf_reqstat.TextMatrix(srno, 5) = rs_allreq("APNTDT") & ""
                             End If
                             
                             msf_reqstat.col = 6
                             If UCase(EXAM) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                             msf_reqstat.TextMatrix(srno, 6) = rs_allreq("EXAMDT") & ""
                             End If
                             
                             
                             msf_reqstat.col = 7
                             If UCase(REPORT) = "YESF" Then
                            msf_reqstat.CellBackColor = vbGreen
                            msf_reqstat.TextMatrix(srno, 7) = "FNL " & rs_allreq("FREPORTDT")
                             End If
                            If UCase(REPORT) = "YEST" Then
                            msf_reqstat.CellBackColor = vbGreen
                            msf_reqstat.TextMatrix(srno, 7) = "PRV " & rs_allreq("PREPORTDT")
                             End If
                             
                              msf_reqstat.col = 8
                            If Mid(rs_allreq("REQNNO"), 4, 2) = "IH" Then
                                
                                seldrequisitn_qr(0) = rs_allreq("REQNNO")
                                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                                If seldrequisitn_rs.RowCount > 0 Then
                                    If Len(Trim(seldrequisitn_rs("SPIHREQNO"))) > 0 Then
                                    
                                         msf_reqstat.CellBackColor = vbGreen
                                         msf_reqstat.TextMatrix(srno, 8) = req_chglen(Trim(seldrequisitn_rs("spihreqno")))
                                     End If
                                End If
                            End If
                            If Mid(rs_allreq("reqnno"), 4, 2) = "SP" Then
                                
                                'SELDREQNOQR1(0) = rs_allreq("REQNNO")
                                'Set SELDREQNORS1 = SELDREQNOQR1.OpenResultset(1, 3)
                                'If SELDREQNORS1.RowCount > 0 Then
                                '    msf_reqstat.CellBackColor = vbGreen
                                '    msf_reqstat.TextMatrix(srno, 8) = req_chglen(SELDREQNORS1("REQNO"))
                                'End If
      ' THE CODE SLICE IS COMMENTED BECAUSE THE REQNO IS NOT REQUIRED BY PENDING REMARKS
      ' FIELD FOR THE SP IS REQUIRED FROM SITEDETAIL TABLE OF DISLIB
                                SELSITEDETRQR1(0) = rs_allreq("REQNNO")
                                Set SELSITEDETRRS = SELSITEDETRQR1.OpenResultset(1, 3)
                                If SELSITEDETRRS.RowCount > 0 Then
                                    msf_reqstat.CellBackColor = vbGreen
                                    msf_reqstat.TextMatrix(srno, 8) = SELSITEDETRRS("FLD1")
                                End If
                            End If
                            msf_reqstat.col = 9
                            seldrequisitn_qr(0) = (rs_allreq("REQNNO"))
                            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                            If seldrequisitn_rs.RowCount > 0 Then
                                msf_reqstat.TextMatrix(srno, 9) = Trim(seldrequisitn_rs("PATHNO"))
                            End If
                             
                                rs_allreq.MoveNext
                               srno = srno + 1
                    Wend
                '    MsgBox msf_reqstat.Rows
                Else
                
                MsgBox "NO DATA is found between these two dates."
                  button_ena_dis ("cancel")
                  clear
                 MaskEd_from_date.Enabled = False
                 MaskEd_to_date.Enabled = False
                 cmd_view.SetFocus
            End If
            
 
End Sub
Private Sub micro_myco()
If Opt_patient.Value = True Then

        CrystalReport1.WindowTitle = "Microbiology Report"
        CrystalReport1.SelectionFormula = "{micro_test.REQNO} = '" & VAR_REQNNO & "'"
        CrystalReport1.Formulas(0) = "namE ='" & Trim(SELTA101_RS("FIRST_NAME")) + " " + Trim(SELTA101_RS("MID_NAME")) + " " + Trim(SELTA101_RS("LAST_NAME")) & "'"
        CrystalReport1.Formulas(1) = "sex ='" & Trim(SELTA101_RS("SEX")) & "'"
        CrystalReport1.Formulas(2) = "age ='" & DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), Date) & "'"

        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\MICRO_MYCO.rpt"

        CrystalReport1.Destination = crptToWindow
        CrystalReport1.WindowShowPrintBtn = True
        CrystalReport1.WindowShowRefreshBtn = True
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1

    Else

        CrystalReport1.WindowTitle = "Microbiology Report"
        CrystalReport1.SelectionFormula = "{micro_test.REQNO} = '" & VAR_REQNNO & "'"
        CrystalReport1.Formulas(0) = "CCnamE ='" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(1) = "sex ='" & Trim(SELEMPBASINFO_RS("GENDER")) & "'"
        CrystalReport1.Formulas(2) = "age ='" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT")) & "'"

        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\MICRO_MYCO_STF.rpt"

        CrystalReport1.Destination = crptToWindow
        CrystalReport1.WindowShowPrintBtn = True
        CrystalReport1.WindowShowRefreshBtn = True
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Action = 1

    End If
End Sub
Private Sub CallToDISCY()
CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCYOLD.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(req_chglen(VAR_REQNNO)) & "'"
        If Trim(REFDR_RS_N("REQNSTAT")) = "T" Then
            str1 = ""
            'STR1 = "Entered By " & Trim(cmb_enteryby.Text)
            CrystalReport1.Formulas(1) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(1) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
        If REFDR_RS_N("MODALITYCD") = "FI" Then
           'If Trim(FIsh_txt_inter.Text) <> "" Then
            'CRYSTALREPORT1.Formulas(10) = "inter = 'abc'" '& STR1 & "'"
           'Else
            CrystalReport1.Formulas(10) = "inter = ''" '& STR1 & "'"
           'End If

           'If Trim(fish_txt_comments.Text) <> "" Then
           ' CRYSTALREPORT1.Formulas(11) = "comments = 'abc'" '& STR1 & "'"
          ' Else
            CrystalReport1.Formulas(11) = "comments = ''" '& STR1 & "'"
          ' End If

          CrystalReport1.Formulas(3) = "PROFIN = '" & "MOLECULAR CYTOGENETICS (FISH) REPORT" & "'"

        Else
          If Trim(txt_Interpre.Text) <> "" Then
            CrystalReport1.Formulas(10) = "inter = 'abc'" '& STR1 & "'"
           Else
            CrystalReport1.Formulas(10) = "inter = ''" '& STR1 & "'"
           End If

           If Trim(txt_comments.Text) <> "" Then
            CrystalReport1.Formulas(11) = "comments = 'abc'" '& STR1 & "'"
           Else
            CrystalReport1.Formulas(11) = "comments = ''" '& STR1 & "'"
           End If
          CrystalReport1.Formulas(3) = "PROFIN = '" & "CONVENTIONAL CYTOGENETICS REPORT" & "'"
        End If
        'CRYSTALREPORT1.Formulas(4) = "cdetails = '" & txt_clinicaldetail.Text & "'"
        'CRYSTALREPORT1.Formulas(5) = "spstatus = '" & txt_specimen.Text & "'"
       ' CRYSTALREPORT1.Formulas(6) = "PROVDIA = '" & txt_provdiag.Text & "'"
      CrystalReport1.Formulas(6) = "NAM='" & txt_name.Text & "'"
      CrystalReport1.WindowState = crptMaximized
      CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

      CrystalReport1.Action = 1
End Sub


Private Sub msf_reqstat_Click()
req_sta = Mid(Trim(msf_reqstat.TextMatrix(msf_reqstat.row, 1)), 1, 16)

Dim pacsaddr1 As String
Dim pacsaddr2 As String
Dim pacsaddr3 As String
Dim pacsaddr As String

If msf_reqstat.col = 8 Then  ' COL 8 IS PICTURE IS ATTEMPTED
  If msf_reqstat.TextMatrix(msf_reqstat.row, 8) = "PICTURE" Then
   pacsaddr1 = "http://centricityweb/ami/html/webviewer.html?viewall&ris_exam_id="
   pacsaddr3 = "&un=TMHPACS&pw=tmhpacs"
   VAR_REQNNO = req_chglen(Left(msf_reqstat.TextMatrix(msf_reqstat.row, 1), 16))   '''req_chglen(mask_edit_req.Text)
   pacsaddr = pacsaddr1 + VAR_REQNNO + pacsaddr3
   ShellExecute GetActiveWindow(), "open", pacsaddr, "", 0&, 1
   Exit Sub
 End If
 Exit Sub
 
End If

Dim str1 As String
Dim str2 As String
Dim STR3 As String
 
DoEvents

msf_reqstat.col = 1
str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
VAR_REQNNO = req_chglen(Left(msf_reqstat.Text, 16))  '''req_chglen(mask_edit_req.Text)
var_caseno = msf_reqstat.TextMatrix(msf_reqstat.row, 2)
selrrequisitnqr(0) = VAR_REQNNO
Set selrrequisitnrs = selrrequisitnqr.OpenResultset(1, 3)
If selrrequisitnrs.RowCount > 0 Then

End If
seldrequisitn_qr(0) = VAR_REQNNO
Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'''DTM BLANK REPORT CHEQUE
If selrrequisitnrs("MODALITYCD") = "SS" Then
    If SELDTMTAB_FLAG = False Then
        SELDTMTAB_QR(0) = VAR_REQNNO
        Set SELDTMTAB_RS = SELDTMTAB_QR.OpenResultset(1, 3)
        SELDTMTAB_FLAG = True
    Else
        SELDTMTAB_QR(0) = VAR_REQNNO
        SELDTMTAB_RS.MoreResults
        SELDTMTAB_RS.Requery
    End If
    'If ((Len(Trim(SELDTMTAB_RS("ABO"))) > 1 And SELDTMTAB_RS("ABO")) <> "") And (Len(Trim(SELDTMTAB_RS("RH"))) > 1 Or Trim(SELDTMTAB_RS("RH")) <> "") And (Len(Trim(SELDTMTAB_RS("ABS"))) > 1 Or Trim(SELDTMTAB_RS("ABS")) <> "") And (Len(Trim(SELDTMTAB_RS("IAT"))) > 1 Or Trim(SELDTMTAB_RS("IAT")) <> "") And (Len(Trim(SELDTMTAB_RS("DAT"))) > 1 Or Trim(SELDTMTAB_RS("DAT")) <> "") And (Len(Trim(SELDTMTAB_RS("HBSAG"))) > 1 Or Trim(SELDTMTAB_RS("HBSAG")) <> "") And (Len(Trim(SELDTMTAB_RS("HCV"))) > 1 Or Trim(SELDTMTAB_RS("HCV")) <> "") And (Len(Trim(SELDTMTAB_RS("HIV"))) > 1 Or Trim(SELDTMTAB_RS("HIV")) <> "") Then
    If Not SELDTMTAB_RS.EOF Then
    If (IsNull(SELDTMTAB_RS("ABO")) Or Trim(SELDTMTAB_RS("ABO")) = "") And (IsNull(SELDTMTAB_RS("RH")) Or Trim(SELDTMTAB_RS("RH")) = "") And (IsNull(Trim(SELDTMTAB_RS("ABS"))) Or Trim(SELDTMTAB_RS("ABS")) = "") And (IsNull(Trim(SELDTMTAB_RS("IAT"))) Or Trim(SELDTMTAB_RS("IAT")) = "") And (IsNull(Trim(SELDTMTAB_RS("DAT"))) Or Trim(SELDTMTAB_RS("DAT")) = "") And (IsNull(Trim(SELDTMTAB_RS("HBSAG"))) Or Trim(SELDTMTAB_RS("HBSAG")) = "") And (IsNull(Trim(SELDTMTAB_RS("HCV"))) Or Trim(SELDTMTAB_RS("HCV")) = "") And (IsNull(Trim(SELDTMTAB_RS("HIV"))) Or Trim(SELDTMTAB_RS("HIV")) = "") Then
        MsgBox " Provisional has not Created !!!", vbInformation, "DTM"
        Exit Sub
    Else
        GoTo DTM
    End If
    End If
End If
'''DTM BLANK REPORT CHEQUE
DTM:
For i = 0 To 30
    CrystalReport1.Formulas(i) = ""
Next i
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'" '& VAR_REQNNO & "'"

If Trim(msf_reqstat.TextMatrix(msf_reqstat.row, 7)) = "" Then
    MsgBox "Provisional Report is Not yet Prepared"
    Exit Sub
End If

If LNK_FLAG = False Then
    LNK_QR(0) = LOGIN
    Set LNK_RS = LNK_QR.OpenResultset(1, 3)
    LNK_FLAG = True
    Else
    LNK_QR(0) = LOGIN
    LNK_RS.MoreResults
    LNK_RS.Requery
End If
If SELTA101_FLAG = False Then
    SELTA101_QR(0) = selrrequisitnrs("CASENO")
    Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
    SELTA101_FLAG = True
Else
    SELTA101_QR(0) = selrrequisitnrs("CASENO")
    SELTA101_RS.MoreResults
    SELTA101_RS.Requery
End If
If SELEMPBASINFOFLAG = False Then
    SELEMPBASINFO_QR(0) = Msk_caseno.Text
    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
    SELEMPBASINFOFLAG = True
Else
    SELEMPBASINFO_QR(0) = Msk_caseno.Text
    SELEMPBASINFO_RS.MoreResults
    SELEMPBASINFO_RS.Requery
End If
selta1012_qr(0) = selrrequisitnrs("CASENO")
Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"


selta1012_qr(0) = selrrequisitnrs("CASENO")
Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)

If selrrequisitnrs("MODALITYCD") = "BF" Then
    BFREPORT
End If
If selrrequisitnrs("MODALITYCD") = "HL" Then
   VAR_REQNNO = req_chglen(Left(msf_reqstat.Text, 16))
    hlreport1
End If
If selrrequisitnrs("MODALITYCD") = "EK" Then
    EKREPORT
End If
'''********BONE MARROW*************
 
'Added by Milind on 03/02/2006 for viewing only committed reports to other hospital user(3rd & 4th If statements)

    If selrrequisitnrs("MODALITYCD") = "BM" Then
    
CrystalReport1.SelectionFormula = "{REPORTBMLF.REQNO} = '" & VAR_REQNNO & "'"
    
      If Trim(LNK_RS("DEPT_CODE")) = "UZZ" Then
MSG:        If SELDBMAREPORT_FLAG = False Then
                SELDBMAREPORT_QR(0) = VAR_REQNNO
                Set SELDBMAREPORT_RS = SELDBMAREPORT_QR.OpenResultset(1, 3)
                SELDBMAREPORT_FLAG = True
            Else
                SELDBMAREPORT_QR(0) = VAR_REQNNO
                SELDBMAREPORT_RS.MoreResults
                SELDBMAREPORT_RS.Requery
            End If
    'amit
    
        If Mid(msf_reqstat.TextMatrix(msf_reqstat.row, 7), 1, 3) = "PRV" Or Mid(msf_reqstat.TextMatrix(msf_reqstat.row, 7), 1, 3) = "FNL" Or Mid(msf_reqstat.TextMatrix(msf_reqstat.row, 7), 1, 3) = "PRO" Then
            If Opt_patient.Value = True Then
            SELDBMASUB_QR(0) = VAR_REQNNO
            Set SELDBMASSUB_RS = SELDBMASUB_QR.OpenResultset(1, 1)
            If SELDBMASSUB_RS.RowCount > 0 Then
                If SELDBMASSUB_RS("FULLREP") = "Y" Then
          
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBM3A.RPT"
                    'Added by Milind on 10-02-06 to view designation of Dr. in Report
                    
                    adndm_qr(0) = VAR_REQNNO
                    Set adndm_rs = adndm_qr.OpenResultset(1, 1)
                    If adndm_rs.RowCount > 0 Then
                        CrystalReport2.ReportFileName = "c:\DISREPORTSEC\DISBM3AADD3.RPT"
                        CrystalReport2.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
                        CrystalReport2.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                        CrystalReport2.Formulas(2) = "PROVDIA = '" & selta1012_rs("PRVDIAGDES") & "'"
                        CrystalReport2.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
                    End If
                    
                    qrfrep_qr(0) = VAR_REQNNO
                    Set rsfrep_rs = qrfrep_qr.OpenResultset(1, 1)
                                        
                    If rsfrep_rs.RowCount > 0 Then
                        If Trim(rsfrep_rs("FREPORTBY")) = "C N NAIR" Then
                            CrystalReport1.Formulas(27) = "desg = 'Officer-Incharge'"
                            CrystalReport2.Formulas(27) = "desg = 'Officer-Incharge'"
                        End If
                        If Trim(rsfrep_rs("FREPORTBY")) = "MONA ANAND" Then
                            CrystalReport1.Formulas(27) = "desg = 'Pathologist'"
                            CrystalReport2.Formulas(27) = "desg = 'Officer-Incharge'"
                        End If
                        If Trim(rsfrep_rs("FREPORTBY")) = "SUMMET GUJRAL" Then
                           CrystalReport1.Formulas(27) = "desg = 'Asst.Officer-Incharge'"
                           CrystalReport2.Formulas(27) = "desg = 'Officer-Incharge'"
                        End If
                    End If
                    'upto here
                End If
          
                If SELDBMASSUB_RS("FULLREP") = "N" Then
          
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBM3AI.RPT"
                    
                    'Added by Milind on 10-02-06 to view designation of Dr. in Report
                    qrfrep_qr(0) = VAR_REQNNO
                    Set rsfrep_rs = qrfrep_qr.OpenResultset(1, 1)
                    
                    If rsfrep_rs.RowCount > 0 Then
                        If Trim(rsfrep_rs("FREPORTBY")) = "C N NAIR" Then
                            CrystalReport1.Formulas(27) = "desg = 'Officer-Incharge'"
                            CrystalReport2.Formulas(27) = "desg = 'Officer-Incharge'"
                        End If
                        If Trim(rsfrep_rs("FREPORTBY")) = "MONA ANAND" Then
                            CrystalReport1.Formulas(27) = "desg = 'Pathologist'"
                            CrystalReport2.Formulas(27) = "desg = 'Officer-Incharge'"
                        End If
                        If Trim(rsfrep_rs("FREPORTBY")) = "SUMMET GUJRAL" Then
                           CrystalReport1.Formulas(27) = "desg = 'Asst.Officer-Incharge'"
                           CrystalReport2.Formulas(27) = "desg = 'Officer-Incharge'"
                        End If
                    End If
                    'upto here
                End If
            End If
       
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            
            If Not selta1012_rs.EOF Then
                CrystalReport1.Formulas(2) = "PROVDIA = '" & selta1012_rs("PRVDIAGDES") & "'"
             
            End If
            
            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
                
                str1 = ""
                str1 = "Entered By " & selrrequisitnrs("ENTERBY")
                If Len(Trim(selrrequisitnrs("ENTERBY"))) > 0 Then
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                    CrystalReport2.Formulas(9) = "ENT = '" & str1 & "'"
                Else
                    CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
                    CrystalReport2.Formulas(9) = "ENT = '" & "" & "'"
                End If
            Else
                str1 = ""
                str1 = "Entered By " & selrrequisitnrs("ENTERBY")
                CrystalReport1.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                CrystalReport2.Formulas(9) = "ENT = '" & str1 & "'"
               
            End If
            
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
            CrystalReport2.Formulas(12) = "USRID = '" & str & "'"
            
            If Trim(SELDBMAREPORT_RS("CD2")) = "" And Trim(SELDBMAREPORT_RS("CD3")) = "" And Trim(SELDBMAREPORT_RS("CD4")) = "" And Trim(SELDBMAREPORT_RS("CD5")) = "" And Trim(SELDBMAREPORT_RS("CD7")) = "" And Trim(SELDBMAREPORT_RS("CD8")) = "" And Trim(SELDBMAREPORT_RS("CYTOCD3")) = "" Then
                CrystalReport1.Formulas(13) = "tmarkers = '" & "" & "'"
            Else
                CrystalReport1.Formulas(13) = "tmarkers = '" & "abc" & "'"
            End If
            If Trim(SELDBMAREPORT_RS("CD10")) = "" And Trim(SELDBMAREPORT_RS("CD19")) = "" And Trim(SELDBMAREPORT_RS("FMC7")) = "" And Trim(SELDBMAREPORT_RS("CD22")) = "" And Trim(SELDBMAREPORT_RS("CD23")) = "" And Trim(SELDBMAREPORT_RS("smig")) = "" And Trim(SELDBMAREPORT_RS("CD79B")) = "" And Trim(SELDBMAREPORT_RS("CD25")) = "" And Trim(SELDBMAREPORT_RS("CD11C")) = "" And Trim(SELDBMAREPORT_RS("cd103")) = "" And Trim(SELDBMAREPORT_RS("kappa")) = "" And Trim(SELDBMAREPORT_RS("lamda")) = "" And Trim(SELDBMAREPORT_RS("DBA44")) = "" And Trim(SELDBMAREPORT_RS("cytoEP22")) = "" And Trim(SELDBMAREPORT_RS("cytocd79a")) = "" And Trim(SELDBMAREPORT_RS("CYTOMU")) = "" Then
                CrystalReport1.Formulas(14) = "bmarkers = '" & "" & "'"
            Else
                CrystalReport1.Formulas(14) = "bmarkers = '" & "abc" & "'"
            End If
            If Trim(SELDBMAREPORT_RS("CD34")) = "" And Trim(SELDBMAREPORT_RS("HLADR")) = "" And Trim(SELDBMAREPORT_RS("CD38")) = "" And Trim(SELDBMAREPORT_RS("TDT")) = "" And Trim(SELDBMAREPORT_RS("CD138")) = "" And Trim(SELDBMAREPORT_RS("ANTIMPO")) = "" Then
                CrystalReport1.Formulas(15) = "others = '" & "" & "'"
            Else
                CrystalReport1.Formulas(15) = "others = '" & "abc" & "'"
            End If
            If Trim(SELDBMAREPORT_RS("cd13")) = "" And Trim(SELDBMAREPORT_RS("Cd14")) = "" And Trim(SELDBMAREPORT_RS("CD15")) = "" And Trim(SELDBMAREPORT_RS("CD33")) = "" And Trim(SELDBMAREPORT_RS("cd36")) = "" And Trim(SELDBMAREPORT_RS("cd4142")) = "" And Trim(SELDBMAREPORT_RS("Glycophoin")) = "" And Trim(SELDBMAREPORT_RS("CD68")) = "" And Trim(SELDBMAREPORT_RS("CD42")) = "" And Trim(SELDBMAREPORT_RS("CD117")) = "" Then
                CrystalReport1.Formulas(16) = "mm = '" & "" & "'"
            Else
                CrystalReport1.Formulas(16) = "mm = '" & "abc" & "'"
            End If
            If Trim(SELDBMAREPORT_RS("IRONSTAIN")) = "" And Trim(SELDBMAREPORT_RS("GIDEROBLAS")) = "" And Trim(SELDBMAREPORT_RS("sudanblack")) = "" And Trim(SELDBMAREPORT_RS("TOLUIDINEB")) = "" Then
                CrystalReport1.Formulas(17) = "cci = '" & "" & "'"
            Else
                CrystalReport1.Formulas(17) = "cci = '" & "abc" & "'"
            End If
            If Trim(SELDBMAREPORT_RS("myelopero")) = "" And Trim(SELDBMAREPORT_RS("nonspester")) = "" And Trim(SELDBMAREPORT_RS("trap")) = "" Then
                CrystalReport1.Formulas(18) = "ccm = '" & "" & "'"
            Else
                CrystalReport1.Formulas(18) = "ccm = '" & "abc" & "'"
            End If
            If Trim(SELDBMAREPORT_RS("IMPRINT")) = "" And Trim(SELDBMAREPORT_RS("abcellblas")) = "" Then
                CrystalReport1.Formulas(19) = "im = '" & "" & "'"
            Else
                CrystalReport1.Formulas(19) = "im = '" & "abc" & "'"
            End If
            If Trim(SELDBMAREPORT_RS("mcellulari")) = "" And Trim(SELDBMAREPORT_RS("morphology")) = "" Then
                CrystalReport1.Formulas(20) = "ms = '" & "" & "'"
            Else
                CrystalReport1.Formulas(20) = "ms = '" & "abc" & "'"
            End If
            If Trim(SELDBMAREPORT_RS("ecellulari")) = "" And Trim(SELDBMAREPORT_RS("maturation")) = "" Then
                CrystalReport1.Formulas(21) = "es = '" & "" & "'"
            Else
                CrystalReport1.Formulas(21) = "es = '" & "abc" & "'"
            End If
            If Trim(SELDBMAREPORT_RS("mature")) = "" And Trim(SELDBMAREPORT_RS("immature")) = "" And Trim(SELDBMAREPORT_RS("atypbinu")) = "" And Trim(SELDBMAREPORT_RS("plasmablas")) = "" Then
                CrystalReport1.Formulas(22) = "pc = '" & "" & "'"
            Else
                CrystalReport1.Formulas(22) = "pc = '" & "abc" & "'"
            End If
            If Trim(SELDBMAREPORT_RS("MLYMPHOCYT")) = "" And Trim(SELDBMAREPORT_RS("ALYMPHOCYT")) = "" And Trim(SELDBMAREPORT_RS("Plasmacell")) = "" Then
                CrystalReport1.Formulas(23) = "ls = '" & "" & "'"
            Else
                CrystalReport1.Formulas(23) = "ls = '" & "abc" & "'"
            End If
            If Trim(SELDBMAREPORT_RS("PROMYELOCY")) = "" And Trim(SELDBMAREPORT_RS("PROMYELOCY")) = "" And Trim(SELDBMAREPORT_RS("MMYELOCYTE")) = "" And Trim(SELDBMAREPORT_RS("polymorphs")) = "" And Trim(SELDBMAREPORT_RS("basophils")) = "" And Trim(SELDBMAREPORT_RS("eosinophil")) = "" And Trim(SELDBMAREPORT_RS("monocytes")) = "" And Trim(SELDBMAREPORT_RS("PROMONOCYT")) = "" Then
                CrystalReport1.Formulas(24) = "mys = '" & "" & "'"
            Else
                CrystalReport1.Formulas(24) = "mys = '" & "abc" & "'"
            End If
            If Trim(SELDBMAREPORT_RS("ECELLULARI")) = "" And Trim(SELDBMAREPORT_RS("MERATIO")) = "" Then
                CrystalReport1.Formulas(25) = "mor = '" & "" & "'"
            Else
                CrystalReport1.Formulas(25) = "mor = '" & "abc" & "'"
            End If
            
        End If
    
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
        
        End If
        If adndm_rs.RowCount > 0 Then
            CrystalReport2.WindowState = crptMaximized
            CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
            CrystalReport2.Action = 1
        End If
    Else
        If Mid(msf_reqstat.TextMatrix(msf_reqstat.row, 7), 1, 3) = "FNL" Then
            SELDBMAREPORT_FLAG = False
            GoTo MSG:
        Else
            MsgBox "Provisional Report Cannot be viewed"
            Exit Sub
        End If
    End If
End If

''******BLOOD PRODUCT ****************

If selrrequisitnrs("MODALITYCD") = "BP" Then
    If SELDTMNO_FLAG = False Then
        SELDTMNO_QR(0) = selrrequisitnrs("CASENO")
        Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
        SELDTMNO_FLAG = True
    Else
        SELDTMNO_QR(0) = selrrequisitnrs("CASENO")
        SELDTMNO_RS.MoreResults
        SELDTMNO_RS.Requery
    End If
    If Opt_patient.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBP.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(1) = "MOD = '" & "BLOOD PRODUCT" & "'"
        If selrrequisitnrs("PATNSTAT") = "OP" Then
            CrystalReport1.Formulas(2) = "ST = '" & selrrequisitnrs("pat_catg") & "/" & " Out Patient" & "'"
        Else
            CrystalReport1.Formulas(2) = "ST = '" & selrrequisitnrs("pat_catg") & "/" & " In patient" & "'"
        End If
        CrystalReport1.Formulas(3) = "SEXAGE = '" & SELTA101_RS("SEX") & "/" & DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), selrrequisitnrs("REQNDT")) & "'"
        If selrrequisitnrs("PATNSTAT") = "IP" Then
            CrystalReport1.Formulas(4) = "WRBED = '" & seldrequisitn_rs("WARD_NAME") & "/" & seldrequisitn_rs("bedno") & "'"
        Else
            CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
        End If
        CrystalReport1.Formulas(5) = "PATHNO = '" & SELDTMNO_RS("DTMNO") & "'"
        CrystalReport1.Formulas(6) = "USRID = '" & str & "'"
    Else
        
    End If
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport1.Action = 1
End If
''******BLOOD PRODUCT ****************
''******SCREENING TEST ****************
If selrrequisitnrs("MODALITYCD") = "SS" Then
    If Opt_patient.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtm.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If seldrequisitn_rs.RowCount > 0 Then
            CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
    Else
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtmreq.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If seldrequisitn_rs.RowCount > 0 Then
            CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
        If SELEMPBASINFO_RS.RowCount > 0 Then
            CrystalReport1.Formulas(3) = "NAM='" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        Else
            CrystalReport1.Formulas(3) = "NAM= '" & " " & "'"
        End If
        CrystalReport1.Formulas(4) = "STDOB='" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(5) = "DOB='" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT")) & "'"
    End If
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport1.Action = 1
End If
''*****SCREENING TEST ****************
''*******HLA TYPING *****************
If selrrequisitnrs("MODALITYCD") = "HT" Then
    If Opt_patient.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISHLA.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        If Len(Trim(seldrequisitn_rs("LABNO"))) > 0 Then
            CrystalReport1.Formulas(8) = "NAM = '" & seldrequisitn_rs("NMETERIALS") & "'"
            CrystalReport1.Formulas(9) = "RELATION = 'Relation' "
        Else
            CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELTA101_RS("FIRST_NAME")) + " " + Trim(SELTA101_RS("MID_NAME")) + " " + Trim(SELTA101_RS("LAST_NAME")) & "'"
            CrystalReport1.Formulas(9) = "RELATION ='" & " " & "'"
        End If
        CrystalReport1.Formulas(9) = "dtmno='" & " " & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISHLAstaff.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        If Len(Trim(seldrequisitn_rs("LABNO"))) > 0 Then
            CrystalReport1.Formulas(8) = "NAM = '" & seldrequisitn_rs("NMETERIALS") & "'"
            CrystalReport1.Formulas(9) = "RELATION = 'Relation' "
        Else
            CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
            CrystalReport1.Formulas(9) = "RELATION ='" & " " & "'"
        End If
        CrystalReport1.Formulas(10) = "STDOB ='" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(11) = "DOB ='" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT")) & "'"
        CrystalReport1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    End If
End If

'********** Shraddha Pandit 11/11/08 ***********
If selrrequisitnrs("MODALITYCD") = "MC" Then

    micro_myco

'    If opt_patient.Value = True Then
'
'        CRYSTALREPORT1.WindowTitle = "Microbiology Report"
'        CRYSTALREPORT1.SelectionFormula = "{micro_test.REQNO} = '" & VAR_REQNNO & "'"
'        CRYSTALREPORT1.Formulas(0) = "nam ='" & Trim(SELTA101_RS("FIRST_NAME")) + " " + Trim(SELTA101_RS("MID_NAME")) + " " + Trim(SELTA101_RS("LAST_NAME")) & "'"
'        CRYSTALREPORT1.Formulas(1) = "sex ='" & Trim(SELTA101_RS("SEX")) & "'"
'        CRYSTALREPORT1.Formulas(2) = "age ='" & DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), Date) & "'"
'
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\MICRO_MYCO.rpt"
'
'        CRYSTALREPORT1.Destination = crptToWindow
'        CRYSTALREPORT1.WindowShowPrintBtn = True
'        CRYSTALREPORT1.WindowShowRefreshBtn = True
'        CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'        CRYSTALREPORT1.WindowState = crptMaximized
'        CRYSTALREPORT1.Action = 1
'
'    Else
'
'        CRYSTALREPORT1.WindowTitle = "Microbiology Report"
'        CRYSTALREPORT1.SelectionFormula = "{micro_test.REQNO} = '" & VAR_REQNNO & "'"
'        CRYSTALREPORT1.Formulas(0) = "nam ='" & Trim(SELTA101_RS("FIRST_NAME")) + " " + Trim(SELTA101_RS("MID_NAME")) + " " + Trim(SELTA101_RS("LAST_NAME")) & "'"
'        CRYSTALREPORT1.Formulas(1) = "sex ='" & Trim(SELTA101_RS("SEX")) & "'"
'        CRYSTALREPORT1.Formulas(2) = "age ='" & DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), Date) & "'"
'
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\MICRO_MYCO_STF.rpt"
'
'        CRYSTALREPORT1.Destination = crptToWindow
'        CRYSTALREPORT1.WindowShowPrintBtn = True
'        CRYSTALREPORT1.WindowShowRefreshBtn = True
'        CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'        CRYSTALREPORT1.WindowState = crptMaximized
'        CRYSTALREPORT1.Action = 1
'
'    End If
End If
'***********************************************

''*******HLA TYPING *****************
 '*** SURGICAL PAHOLOGY REPORT
 '*** IMMUNIOHISTO CHEMISTRY REPORT
If selrrequisitnrs("MODALITYCD") = "IH" Then
    If Opt_patient.Value = True Then
        If Len(Trim(seldrequisitn_rs("SPIHREQNO"))) = 0 Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISIH.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If selrrequisitnrs.RowCount > 0 Then
               If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHC REPORT" & "'"
                Else
                    CrystalReport1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHC REPORT" & "'"
                End If
            End If
            If selta1012_rs.RowCount > 0 Then
                CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
            Else
                CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
            End If
            If seldrequisitn_rs.RowCount > 0 Then
                CrystalReport1.Formulas(3) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
            Else
                CrystalReport1.Formulas(3) = "PREPATH = '" & "" & "'"
            End If
            CrystalReport1.Formulas(4) = "USRID = '" & str & "'"
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
            CrystalReport1.Action = 1
        End If
    Else
        If Len(Trim(seldrequisitn_rs("SPIHREQNO"))) = 0 Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISIHSTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If selrrequisitnrs.RowCount > 0 Then
               If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(1) = "PROFIN = '" & "PROVISIONAL HISTOPATH & IHC REPORT" & "'"
                Else
                    CrystalReport1.Formulas(1) = "PROFIN = '" & "FINAL HISTOPATH & IHC REPORT" & "'"
                End If
            End If
            If selta1012_rs.RowCount > 0 Then
                CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
            Else
                CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
            End If
            If seldrequisitn_rs.RowCount > 0 Then
                CrystalReport1.Formulas(3) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
            Else
                CrystalReport1.Formulas(3) = "PREPATH = '" & "" & "'"
            End If
            CrystalReport1.Formulas(4) = "USRID = '" & str & "'"
            CrystalReport1.Formulas(5) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
            CrystalReport1.Formulas(6) = "SEX = '" & SELEMPBASINFO_RS("GENDER") & "'"
            CrystalReport1.Formulas(4) = "SIT = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
            CrystalReport1.Action = 1
        End If
    End If
  ' *** IMMUNIO HISTO CHEMISTRY REPORT
End If
 
If selrrequisitnrs("MODALITYCD") = "SP" Then
    If Not LNK_RS.EOF Then
    If Trim(LNK_RS("DEPT_CODE")) = "FZZ" Then
MSG1:   For i = 0 To 25
            CrystalReport1.Formulas(i) = ""
        Next i
        ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Opt_patient.Value = True Then
                    
            If selrrequisitnrs("COMMITFLAG") = "Y" Then
                SELDADDENDUMQR(0) = VAR_REQNNO
                Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
                If SELDADDENDUMRS.RowCount > 0 Then
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
                Else
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
                End If
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
            End If
                seldrequisitn_qr(0) = VAR_REQNNO
                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                If seldrequisitn_rs.RowCount > 0 Then
                    CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
                Else
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                End If
                selta1012_qr(0) = selrrequisitnrs("CASENO")
                Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                If selta1012_rs.RowCount > 0 Then
                    CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
                Else
                    CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
                End If
                    
                Dim z
                Dim Y
                
                If Not seldrequisitn_rs("SITE") <> "" Then
                    SELMA301_QR(0) = seldrequisitn_rs("SITE")
                Else
                    SELMA301_QR(0) = " "
                End If
                    
                Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
                If SELMA301_RS.RowCount > 0 Then
                    CrystalReport1.Formulas(3) = "SIT = '" & Trim(SELMA301_RS("SITEDESC")) & "'"
                Else
                    CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                    If Trim(seldrequisitn_rs("OPPROCEDUR")) <> "" Then
                        z = InStr(1, Trim(seldrequisitn_rs("OPPROCEDUR")), "'")
                        If z > 0 Then
                            Y = Trim(seldrequisitn_rs("OPPROCEDUR"))
                            u = Len(Y)
                            X = Left(Y, z - 1)
                            w = Right(Y, u - (z))
                            Y = X & w
                            CrystalReport1.Formulas(4) = "OPPRO = '" & Y & "'"
                        Else
                            CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(seldrequisitn_rs("OPPROCEDUR")) & "'"
                        End If
                    Else
                        CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(seldrequisitn_rs("OPPROCEDUR")) & "'"
                    End If
                Else
                    CrystalReport1.Formulas(4) = "OPPRO = '" & "" & "'"
                End If
                str1 = ""
                If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                    'CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL HISTOPATHOLOGY REPORT" & "'"
                    'Above statement commented and next added by Ms.Shraddha Pandit on 22/07/09 as told by Mrs.Surekha
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
                    str1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                Else
                    'CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL HISTOPATHOLOGY REPORT" & "'"
                    'Above statement commented and next added by Ms.Shraddha Pandit on 22/07/09 as told by Mrs.Surekha
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
                    str1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                End If
                    
                SELDFREPORT_QR(0) = VAR_REQNNO
                Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
                If SELDFREPORT_RS.RowCount > 0 Then
                    If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
                        CrystalReport1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
                    Else
                        CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
                    End If
                    If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
                        CrystalReport1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
                    Else
                        CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
                    End If
                    If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
                        CrystalReport1.Formulas(10) = "GROSDESC = '" & "Gross Description" & "'"
                    Else
                        CrystalReport1.Formulas(10) = "GROSDESC = '" & "" & "'"
                    End If
                    If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
                        CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
                    Else
                        CrystalReport1.Formulas(11) = "MICRODESC = '" & "" & "'"
                    End If
                    If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
                        CrystalReport1.Formulas(12) = "IMPN = '" & "Impression" & "'"
                    Else
                        CrystalReport1.Formulas(12) = "IMPN = '" & "" & "'"
                    End If
                     
                End If
                    
                CrystalReport1.Formulas(8) = "USRID = '" & str & "'"
                    
                'CrystalReport1.Formulas(10) = "GROSDESC = '" & "Gross description" & "'"
                'CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            Else
                If selrrequisitnrs("COMMITFLAG") = "Y" Then
                    SELDADDENDUMQR(0) = VAR_REQNNO
                    Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
                    If SELDADDENDUMRS.RowCount > 0 Then
                        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADDSTAFF.RPT"
                    Else
                        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
                    End If
                Else
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENS.RPT"
                End If
                seldrequisitn_qr(0) = VAR_REQNNO
                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                If seldrequisitn_rs.RowCount > 0 Then
                    CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
                Else
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                End If
                    
                SELMA301_QR(0) = seldrequisitn_rs("SITE")
                Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
                If SELMA301_RS.RowCount > 0 Then
                    CrystalReport1.Formulas(3) = "SIT = '" & Trim(SELMA301_RS("SITEDESC")) & "'"
                Else
                    CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                    CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(seldrequisitn_rs("OPPROCEDUR")) & "'"
                Else
                    CrystalReport1.Formulas(4) = "OPPRO = '" & "" & "'"
                End If
                str1 = ""
                If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL HISTOPATHOLOGY REPORT" & "'"
                    str1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                Else
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL HISTOPATHOLOGY REPORT" & "'"
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                End If
                SELDFREPORT_QR(0) = VAR_REQNNO
                Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
                If SELDFREPORT_RS.RowCount > 0 Then
                    If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
                        CrystalReport1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
                    Else
                        CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
                    End If
                    If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
                        CrystalReport1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
                    Else
                        CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
                    End If
                    If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
                        CrystalReport1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
                    Else
                        CrystalReport1.Formulas(10) = "GROSDESC = '" & "" & "'"
                    End If
                    If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
                        CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
                    Else
                        CrystalReport1.Formulas(11) = "MICRODESC = '" & "" & "'"
                    End If
                    If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
                        CrystalReport1.Formulas(12) = "IMPN = '" & "Impression" & "'"
                    Else
                        CrystalReport1.Formulas(12) = "IMPN = '" & "" & "'"
                    End If

                End If
                    
                CrystalReport1.Formulas(8) = "USRID = '" & str & "'"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
            End If
        Else
            'If Mid(msf_reqstat.TextMatrix(msf_reqstat.row, 7), 1, 3) = "FNL" Then
            'Above statement commented and next added by Shraddha on 23/07/09 as told by Mrs.Surekha
            If Mid(msf_reqstat.TextMatrix(msf_reqstat.row, 7), 1, 3) = "FNL" Or Mid(msf_reqstat.TextMatrix(msf_reqstat.row, 7), 1, 3) = "HPR" Then
                'SELDBMAREPORT_FLAG = False
                GoTo MSG1:
            Else
                MsgBox "Provisional Report Cannot be viewed"
                Exit Sub
            End If
        End If
 ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
    End If
    End If
 '**** FNAC REPORT
    If selrrequisitnrs("MODALITYCD") = "FA" Then
        If Trim(LNK_RS("DEPT_CODE")) = "FZZ" Then
MSG2:       If Opt_patient.Value = True Then
                If selrrequisitnrs("COMMITFLAG") = "Y" Then
                
                    SELDADDENDUMFAQR(0) = VAR_REQNNO
                    Set SELDADDENDUMFARS = SELDADDENDUMFAQR.OpenResultset(1, 3)
                    
                    If SELDADDENDUMFARS.RowCount > 0 Then
                        
                        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADD.RPT"
                        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                        str1 = ""
                        str1 = "Entered By " & selrrequisitnrs("ENTERBY")
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT WITH ADDENDUM" & "'"
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
                        
                    Else
                        
                        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFA.RPT"
                        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                    
                        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                        
                            CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                            str1 = ""
                            str1 = "Entered By " & selrrequisitnrs("ENTERBY")
                            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                        
                        Else
                        
                            str1 = ""
                            str1 = "Entered By " & selrrequisitnrs("ENTERBY")
                            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                        
                        End If
                    
                        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
                    
                    End If
                                        
                Else
                        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFA.RPT"
                        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                        str1 = ""
                        str1 = "Entered By " & selrrequisitnrs("ENTERBY")
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
                End If
                                     
                SELDFN_QR(0) = req_chglen(VAR_REQNNO)
                Set SELDFN_RS = SELDFN_QR.OpenResultset(1, 1)
                    
                If IsNull(SELDFN_RS("MICREPORT")) Then
                    CrystalReport1.Formulas(10) = "fn ='" & " " & "'"
                Else
                    CrystalReport1.Formulas(10) = "fn ='" & "12" & "'"
                End If
                    
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        
                CrystalReport1.Action = 1
            Else
                If (selrrequisitnrs("REQNSTAT") = "T" And selrrequisitnrs("COMMITFLAG") = "") Or (selrrequisitnrs("REQNSTAT") = "F" And selrrequisitnrs("COMMITFLAG") = "Y") Then
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFASTAFF.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                    If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                        str1 = ""
                        str1 = "Entered By " & selrrequisitnrs("ENTERBY")
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                     Else
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                    End If
                    CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
                Else
                    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADDSTAFF.RPT"
                    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                        str1 = ""
                        str1 = "Entered By " & selrrequisitnrs("ENTERBY")
                        CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT WITH ADDENDUM" & "'"
                        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
                End If
                CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
                CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
                CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        
                CrystalReport1.Action = 1
            End If
            
        Else
            'If Mid(msf_reqstat.TextMatrix(msf_reqstat.row, 7), 1, 3) = "FNL" Then
            'Above statement commented and next added by Shraddha on 04/08/09 as told by Mrs.Surekha
            If Mid(msf_reqstat.TextMatrix(msf_reqstat.row, 7), 1, 3) = "FNL" Or Mid(msf_reqstat.TextMatrix(msf_reqstat.row, 7), 1, 3) = "FAR" Then
            'SELDBMAREPORT_FLAG = False
            GoTo MSG2:
            Else
                MsgBox "Provisional Report Cannot be viewed"
                Exit Sub
            End If
            
        End If
    End If
 '**** FNAC REPORT
 '**** CYTOLOGY REPORT
    If selrrequisitnrs("MODALITYCD") = "CG" Or selrrequisitnrs("MODALITYCD") = "CN" Then
        If Opt_patient.Value = True Then
            If Trim(selrrequisitnrs("REQNSTAT")) <> "F" Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCG.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                '            If Trim(TXT_CSITE.Text) = "" Then
                '                CrystalReport1.Formulas(1) = "STL = '" & "" & "'"
                '            Else
                '                CrystalReport1.Formulas(1) = "STL = '" & "Site  " & TXT_CSITE.Text & "'"
                '            End If
                If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
                    str1 = ""
                    str1 = "Entered By " & Trim(selrrequisitnrs("ENTERBY"))
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                Else
                    str1 = ""
                    str1 = "Entered By " & Trim(selrrequisitnrs("ENTERBY"))
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                End If
                CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(12) = "NOSLIDES=" & "" & seldrequisitn_rs("LABNO")
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADD.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(12) = "NOSLIDES = " & "" & seldrequisitn_rs("LABNO")
            End If
            If Trim(selrrequisitnrs("PREPORTBY")) <> "" And Trim(selrrequisitnrs("FREPORTBY")) <> "" And UCase(Trim(selrrequisitnrs("PREPORTBY"))) <> UCase(Trim(selrrequisitnrs("FREPORTBY"))) Then
                CrystalReport1.Formulas(11) = "P= '" & Trim(selrrequisitnrs("PREPORTBY")) & "'"
                CrystalReport1.Formulas(12) = "F= '" & Trim(selrrequisitnrs("FREPORTBY")) & "'"
                CrystalReport1.Formulas(13) = "PCAP= '" & "Cytotechnologist" & " '"
                CrystalReport1.Formulas(14) = "FCAP= '" & "Pathologist" & " '"
                CrystalReport1.Formulas(16) = "SIGNNAME= '" & Trim(selrrequisitnrs("FREPORTBY")) & " '"
            End If
            If Trim(selrrequisitnrs("PREPORTBY")) <> "" And Trim(selrrequisitnrs("FREPORTBY")) = "" Then
                CrystalReport1.Formulas(11) = "PONLY= '" & Trim(selrrequisitnrs("PREPORTBY")) & "'"
                CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
                CrystalReport1.Formulas(16) = "SIGNNAME= '" & Trim(selrrequisitnrs("PREPORTBY")) & " '"
            End If
            If Trim(selrrequisitnrs("PREPORTBY")) <> "" And Trim(selrrequisitnrs("FREPORTBY")) <> "" And UCase(Trim(selrrequisitnrs("PREPORTBY"))) = UCase(Trim(selrrequisitnrs("FREPORTBY"))) Then
                CrystalReport1.Formulas(11) = "PONLY= '" & Trim(selrrequisitnrs("PREPORTBY")) & "'"
                CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
                                CrystalReport1.Formulas(16) = "SIGNNAME= '" & Trim(selrrequisitnrs("PREPORTBY")) & " '"
            End If
            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
             CrystalReport1.Formulas(16) = "SIGNNAME= '" & "" & " '"
             End If
             
        Else
            If selrrequisitnrs("REQNSTAT") <> "F" Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGSTAFF.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL CYTOLOGY REPORT" & "'"
                    str1 = ""
                    str1 = "Entered By " & Trim(selrrequisitnrs("ENTERBY"))
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                Else
                    str1 = ""
                    str1 = "Entered By " & Trim(selrrequisitnrs("ENTERBY"))
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL CYTOLOGY REPORT" & "'"
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                End If
                CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADDSTAFF.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
            End If
            CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
            CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
            CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
        End If
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

        CrystalReport1.Action = 1
    End If
 '**** CYTOLOGY REPORT
 '**** BIOCHEMISTRY REPORT
    If selrrequisitnrs("MODALITYCD") = "BC" Then
        If Opt_patient.Value = True Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBCUNI.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(prepath) = "" Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
            End If
            'CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & selrrequisitnrs("enterby")
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBCSUNI.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & selrrequisitnrs("enterby")
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
'        CRYSTALREPORT1.Formulas(13) = "PAC = '" & "" & "'"
'
    CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
    CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
    'CRYSTALREPORT1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
    End If
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport1.Action = 1
 End If
 '**** BIOCHEMISTRY REPORT
 '**** EMERGENCY LAB REPORT
 If selrrequisitnrs("MODALITYCD") = "EL" Then
    If Opt_patient.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEL.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(prepath) = "" Then
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & selrrequisitnrs("ENTERBY")
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        If pcount <> 0 Then
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        End If
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISELSTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(prepath) = "" Then
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & selrrequisitnrs("ENTERBY")
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        If pcount <> 0 Then
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        End If
         CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
    End If
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
 End If
 '**** EMERGENCY LAB REPORT
 '**** TUMOUR MARKER REPORT
 If selrrequisitnrs("MODALITYCD") = "TM" Then
     If DMULPROCEFLAG = False Then
         DMULPROCE_QR(0) = VAR_REQNNO
         Set DMULPROCE_RS = DMULPROCE_QR.OpenResultset(1, 3)
         DMULPROCEFLAG = True
     Else
         DMULPROCE_QR(0) = VAR_REQNNO
         DMULPROCE_RS.MoreResults
         DMULPROCE_RS.Requery
     End If
     If DMULPROCE_RS.RowCount > 0 Then
         While Not DMULPROCE_RS.EOF
             If Trim(DMULPROCE_RS("PRO_CODE")) = "F816" Then
                 IMFLAG = True
             Else
                 IMFLAG1 = True
             End If
             DMULPROCE_RS.MoveNext
         Wend
     End If
     
     '***** Shraddha 28/07/10 *****
    str = "SELECT * FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
    Set VWDMULPROCE_QR = rdb.CreateQuery("", str)

    VWDMULPROCE_QR(0) = req_chglen(Trim(cbo_req.Text))
    Set VWDMULPROCE_RS = VWDMULPROCE_QR.OpenResultset(1, 3)
    
    If Not VWDMULPROCE_RS.EOF Then
        IMFLAG1 = True 'tm
        IMFLAG = False 'tm1
    Else
        IMFLAG = True 'Tm1
        IMFLAG1 = False 'tm
    End If
     '*****************************
     
     If Opt_patient.Value = True Then
         If IMFLAG1 = True Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM.RPT"
            'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM1.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(prepath) = "" Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
            End If
            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                str1 = ""
                str1 = "Entered by " & selrrequisitnrs("ENTERBY")
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
            CLPTM
            CrystalReport1.Formulas(20) = "COLLAPS TIME = '" & CLPTIME & "'"
        End If
        If IMFLAG = True Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM1.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(prepath) = "" Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
            End If
            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                str1 = ""
                str1 = "Entered by " & selrrequisitnrs("ENTERBY")
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        End If
    Else
        If IMFLAG1 = True Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTMSTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(prepath) = "" Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
            End If
            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                str1 = ""
                str1 = "Entered by " & selrrequisitnrs("ENTERBY")
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
            CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
            CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
            CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
            
            CLPTM
            CrystalReport1.Formulas(17) = "COLLAPS TIME = '" & CLPTIME & "'"
        End If
        If IMFLAG = True Then
            'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM1STAFF.RPT"
            'Added by Shraddha on 13/09/2011
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTMS.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(prepath) = "" Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
            End If
            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                str1 = ""
                str1 = "Entered by " & selrrequisitnrs("ENTERBY")
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(12) = "USRID = '" & ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & "'"
            CrystalReport1.Formulas(15) = "FRSEC = ''"
            CrystalReport1.Formulas(16) = "PROVDIA = ''"
            'CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
            'CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
            CrystalReport1.Formulas(11) = "D0B = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
        End If
    End If
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
End If
'
 IMFLAG = False
 IMFLAG1 = False
 
 '**** TUMOUR MARKER REPORT
 '****ecg reports
If selrrequisitnrs("MODALITYCD") = "PF" Then
    If Opt_patient.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISPF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
    '        If Trim(prepath) = "" Then
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
    '        End If
        selta1012_qr(0) = selrrequisitnrs("CASENO")
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            str1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISPFSTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
    '        If Trim(prepath) = "" Then
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
    '        End If
        selta1012_qr(0) = selrrequisitnrs("CASENO")
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        If Not selta1012_rs.EOF Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        Else
            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
        End If
        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            str1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
    End If
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
End If
If selrrequisitnrs("MODALITYCD") = "ST" Then
    If Opt_patient.Value = True Then
       CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISST.RPT"
       CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
    '        If Trim(prepath) = "" Then
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
    '        End If
       CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
       If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
           CrystalReport1.Formulas(5) = "PROFIN = '" & "STRESS TEST REPORT" & "'"
           str1 = ""
           str1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
           If Len(Trim(selrrequisitnrs("enterby"))) > 0 Then
               CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
           Else
               CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
           End If
           CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
       Else
           CrystalReport1.Formulas(5) = "PROFIN = '" & "STRESS TEST REPORT" & "'"
           CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
       End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        SELSTREPORTQR(0) = req_chglen(VAR_REQNNO)
        Set SELSTREPORTRS = SELSTREPORTQR.OpenResultset(1, 3)
        
        If Len(Trim(SELSTREPORTRS("RISKFACTOR"))) > 0 Then
            CrystalReport1.Formulas(13) = "rk = '" & "Risk Factor" & "'"
        Else
            CrystalReport1.Formulas(13) = "rk = '" & "" & "'"
        End If
        If Len(Trim(SELSTREPORTRS("MAXHR"))) > 0 Then
            CrystalReport1.Formulas(14) = "maxhr = '" & "Max HR" & "'"
        Else
            CrystalReport1.Formulas(14) = "maxhr = '" & "" & "'"
        End If
        If Len(Trim(SELSTREPORTRS("MAXBP"))) > 0 Then
            CrystalReport1.Formulas(15) = "maxBP = '" & "Max BP" & "'"
        Else
            CrystalReport1.Formulas(15) = "maxBP = '" & "" & "'"
        End If
        If Len(Trim(SELSTREPORTRS("DOBPRODUCT"))) > 0 Then
            CrystalReport1.Formulas(16) = "DP = '" & "Double Product" & "'"
        Else
            CrystalReport1.Formulas(16) = "DP = '" & "" & "'"
        End If
        If Len(Trim(SELSTREPORTRS("THRACHIVED"))) > 0 Then
            CrystalReport1.Formulas(17) = "thra = '" & "% of THR Achieved" & "'"
        Else
            CrystalReport1.Formulas(17) = "thra = '" & "" & "'"
        End If
        If Len(Trim(SELSTREPORTRS("STTCHANGES"))) > 0 Then
            CrystalReport1.Formulas(18) = "sttc = '" & "ST/T Changes" & "'"
        Else
            CrystalReport1.Formulas(18) = "sttc = '" & "" & "'"
        End If
        If Len(Trim(SELSTREPORTRS("ARRYTHMIA"))) > 0 Then
            CrystalReport1.Formulas(19) = "arrythmia = '" & "Arrythmia" & "'"
        Else
            CrystalReport1.Formulas(19) = "arrythmia = '" & "" & "'"
        End If
        If Len(Trim(SELSTREPORTRS("METSACHVED"))) > 0 Then
            CrystalReport1.Formulas(20) = "metsa = '" & "METS Achieved" & "'"
        Else
            CrystalReport1.Formulas(20) = "metsa = '" & "" & "'"
        End If
        If Len(Trim(SELSTREPORTRS("IMPRESSION"))) > 0 Then
            CrystalReport1.Formulas(21) = "rt = '" & "Reason for termination" & "'"
        Else
            CrystalReport1.Formulas(21) = "rt = '" & "" & "'"
        End If
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISSTSTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
    '        If Trim(prepath) = "" Then
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
    '        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
           CrystalReport1.Formulas(5) = "PROFIN = '" & "STRESS TEST REPORT" & "'"
           str1 = ""
           str1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
           If Len(Trim(selrrequisitnrs("enterby"))) > 0 Then
               CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
           Else
               CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
           End If
           CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
           CrystalReport1.Formulas(5) = "PROFIN = '" & "STRESS TEST REPORT" & "'"
           CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        SELSTREPORTQR(0) = req_chglen(VAR_REQNNO)
        Set SELSTREPORTRS = SELSTREPORTQR.OpenResultset(1, 3)
        If Not SELSTREPORTRS.EOF Then
            If Len(Trim(SELSTREPORTRS("RISKFACTOR"))) > 0 Then
                CrystalReport1.Formulas(13) = "rk = '" & "Risk Factor" & "'"
            Else
                CrystalReport1.Formulas(13) = "rk = '" & "" & "'"
            End If
        Else
            CrystalReport1.Formulas(13) = "rk = '" & "" & "'"
        End If
        If Not SELSTREPORTRS.EOF Then
            If Len(Trim(SELSTREPORTRS("MAXHR"))) > 0 Then
                CrystalReport1.Formulas(14) = "maxhr = '" & "Max HR" & "'"
            Else
                CrystalReport1.Formulas(14) = "maxhr = '" & "" & "'"
            End If
        Else
            CrystalReport1.Formulas(14) = "maxhr = '" & "" & "'"
        End If
        If Not SELSTREPORTRS.EOF Then
            If Len(Trim(SELSTREPORTRS("MAXBP"))) > 0 Then
                CrystalReport1.Formulas(15) = "maxBP = '" & "Max BP" & "'"
            Else
                CrystalReport1.Formulas(15) = "maxBP = '" & "" & "'"
            End If
        Else
            CrystalReport1.Formulas(15) = "maxBP = '" & "" & "'"
        End If
        If Not SELSTREPORTRS.EOF Then
            If Len(Trim(SELSTREPORTRS("DOBPRODUCT"))) > 0 Then
                CrystalReport1.Formulas(16) = "DP = '" & "Double Product" & "'"
            Else
                CrystalReport1.Formulas(16) = "DP = '" & "" & "'"
            End If
        Else
            CrystalReport1.Formulas(16) = "DP = '" & "" & "'"
        End If
        If Not SELSTREPORTRS.EOF Then
            If Len(Trim(SELSTREPORTRS("THRACHIVED"))) > 0 Then
                CrystalReport1.Formulas(17) = "thra = '" & "% of THR Achieved" & "'"
            Else
                CrystalReport1.Formulas(17) = "thra = '" & "" & "'"
            End If
        Else
            CrystalReport1.Formulas(17) = "thra = '" & "" & "'"
        End If
        If Not SELSTREPORTRS.EOF Then
            If Len(Trim(SELSTREPORTRS("STTCHANGES"))) > 0 Then
                CrystalReport1.Formulas(18) = "sttc = '" & "ST/T Changes" & "'"
            Else
                CrystalReport1.Formulas(18) = "sttc = '" & "" & "'"
            End If
        Else
            CrystalReport1.Formulas(18) = "sttc = '" & "" & "'"
        End If
        If Not SELSTREPORTRS.EOF Then
            If Len(Trim(SELSTREPORTRS("ARRYTHMIA"))) > 0 Then
                CrystalReport1.Formulas(19) = "arrythmia = '" & "Arrythmia" & "'"
            Else
                CrystalReport1.Formulas(19) = "arrythmia = '" & "" & "'"
            End If
        Else
            CrystalReport1.Formulas(19) = "arrythmia = '" & "" & "'"
        End If
        If Not SELSTREPORTRS.EOF Then
            If Len(Trim(SELSTREPORTRS("METSACHVED"))) > 0 Then
                CrystalReport1.Formulas(20) = "metsa = '" & "METS Achieved" & "'"
            Else
                CrystalReport1.Formulas(20) = "metsa = '" & "" & "'"
            End If
        Else
            CrystalReport1.Formulas(20) = "metsa = '" & "" & "'"
        End If
        If Not SELSTREPORTRS.EOF Then
            If Len(Trim(SELSTREPORTRS("IMPRESSION"))) > 0 Then
                CrystalReport1.Formulas(21) = "rt = '" & "Reason for termination" & "'"
            Else
                CrystalReport1.Formulas(21) = "rt = '" & "" & "'"
            End If
        Else
            CrystalReport1.Formulas(21) = "rt = '" & "" & "'"
        End If
        CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
    End If
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
End If
If selrrequisitnrs("MODALITYCD") = "FI" Or selrrequisitnrs("MODALITYCD") = "CC" Then
     ' CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCY.RPT"
     '**************************
    
    
    
    REFDR_QR_N(0) = Msk_caseno.FormattedText
    REFDR_QR_N(1) = VAR_REQNNO
    Set REFDR_RS_N = REFDR_QR_N.OpenResultset(1, 3)
    '*******************************
    CallToDISCY
      
End If


 '****ecg report
 '**** RIS REPORTS
If selrrequisitnrs("MODALITYCD") = "CT" Or selrrequisitnrs("MODALITYCD") = "US" Or selrrequisitnrs("MODALITYCD") = "MM" Or selrrequisitnrs("MODALITYCD") = "MR" Or selrrequisitnrs("MODALITYCD") = "CR" Then
    rq_grid_det(0) = Trim(VAR_REQNNO)
    rs_grid_det.MoreResults
    rs_grid_det.Requery
    If rs_grid_det.RowCount > 0 And (Trim(rs_grid_det("REQNSTAT")) = "T" Or Trim(rs_grid_det("REQNSTAT")) = "F") Then
        If Opt_patient.Value = True Then
            CrystalReport1.Formulas(0) = "provtyp= '" & Trim(rs_grid_det("ENTERBY")) & "'"
            'cr_rep.Formulas(1) = "reporttext= '" & Trim(txt_report.Text) & "'"
            'cr_rep.Formulas(0) = "impresstext= '" & Trim(txt_impression.Text) & "'"
            'cr_rep.Formulas(1) = "finalrep= '" & Trim(txt_finalreport.Text) & "'"
        
            SELRMODALITY_QR(0) = Mid(VAR_REQNNO, 4, 2)
            SELRMODALITY_RS.MoreResults
            SELRMODALITY_RS.Requery
            If SELRMODALITY_RS.RowCount > 0 Then
                CrystalReport1.Formulas(1) = "MODDES='" & SELRMODALITY_RS("MODLITYDES") & "'"
            Else
                CrystalReport1.Formulas(1) = "MODDES='" & "" & "'"
            End If
            CrystalReport1.Formulas(2) = "deptname= '" & "Department of Radiodiagnosis" & "'"
            CrystalReport1.SelectionFormula = "{RISREPORTS.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
            'cr_rep.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
            ''cr_rep.Formulas(0) = "provtypist = '" & txt_type1.Text & "'"
            str1 = ip
            str2 = DATE11
            STR3 = time1
            str1 = str1 & Space(5) & str2 & Space(5) & STR3
            str1 = Space(3) & str1
            CrystalReport1.Formulas(3) = "us_id = '" & str1 & "'"
            CrystalReport1.Formulas(4) = "INOUT = '" & Trim(rs_grid_det("PATNSTAT")) & "'"
            CrystalReport1.Formulas(5) = "reqno = '" & Left(msf_reqstat.Text, 16) & "'"
            If Trim(rs_grid_det("REQNSTAT")) = "T" Or Trim(rs_grid_det("REQNSTAT")) = "F" Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep1.rpt"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep2.rpt"
            End If
        Else
            CrystalReport1.Formulas(0) = "provtyp= '" & Trim(rs_grid_det("ENTERBY")) & "'"
            'cr_rep.Formulas(1) = "reporttext= '" & Trim(txt_report.Text) & "'"
            'cr_rep.Formulas(0) = "impresstext= '" & Trim(txt_impression.Text) & "'"
            'cr_rep.Formulas(1) = "finalrep= '" & Trim(txt_finalreport.Text) & "'"
        
            SELRMODALITY_QR(0) = Mid(VAR_REQNNO, 4, 2)
            SELRMODALITY_RS.MoreResults
            SELRMODALITY_RS.Requery
            If SELRMODALITY_RS.RowCount > 0 Then
                CrystalReport1.Formulas(1) = "MODDES='" & SELRMODALITY_RS("MODLITYDES") & "'"
            Else
                CrystalReport1.Formulas(1) = "MODDES='" & "" & "'"
            End If
            CrystalReport1.Formulas(2) = "deptname= '" & "Department of Radiodiagnosis" & "'"
            CrystalReport1.SelectionFormula = "{RISREPORTS.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
            'cr_rep.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
            ''cr_rep.Formulas(0) = "provtypist = '" & txt_type1.Text & "'"
            str1 = ip
            str2 = DATE11
            STR3 = time1
            str1 = str1 & Space(5) & str2 & Space(5) & STR3
            str1 = Space(3) & str1
            CrystalReport1.Formulas(3) = "us_id = '" & str1 & "'"
            CrystalReport1.Formulas(4) = "INOUT = '" & Trim(rs_grid_det("PATNSTAT")) & "'"
            CrystalReport1.Formulas(5) = "reqno = '" & Left(msf_reqstat.Text, 16) & "'"
            CrystalReport1.Formulas(6) = "PATNAME = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
            CrystalReport1.Formulas(7) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
            CrystalReport1.Formulas(5) = "PATAGE = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
            If Trim(rs_grid_det("REQNSTAT")) = "T" Or Trim(rs_grid_det("REQNSTAT")) = "F" Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep1STAFF.rpt"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep2STAFF.rpt"
            End If
        End If
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

        CrystalReport1.Action = 1
        Exit Sub
    Else
        Exit Sub
    End If
End If
 '**** RIS REPORTS
 '*************************************************************************
 '          wrote for microbiology ****************************************
 ' on 12/12/2002
 
 'If selrrequisitnrs("MODALITYCD") = "BA" Or selrrequisitnrs("MODALITYCD") = "MB" Or selrrequisitnrs("MODALITYCD") = "XX" Or selrrequisitnrs("MODALITYCD") = "SE" Or selrrequisitnrs("MODALITYCD") = "CM" Then
 'Commented and added by Shraddha Pandit on 07/09/09 (added condition for MO)
 If selrrequisitnrs("MODALITYCD") = "BA" Or selrrequisitnrs("MODALITYCD") = "MB" Or selrrequisitnrs("MODALITYCD") = "XX" Or selrrequisitnrs("MODALITYCD") = "SE" Or selrrequisitnrs("MODALITYCD") = "CM" Or selrrequisitnrs("MODALITYCD") = "MO" Then
    QrySts = True 'This flag added by Shraddha on 07/09/09
    Microbiology selrrequisitnrs("MODALITYCD")
 ElseIf selrrequisitnrs("MODALITYCD") = "2D" Then
    Echocardiology selrrequisitnrs("MODALITYCD")
 ElseIf selrrequisitnrs("MODALITYCD") = "ED" Then
    Endoscopy selrrequisitnrs("MODALITYCD")
 End If
 'Exit Sub
'MSG:
'MsgBox Err.Description
End Sub

Private Sub msf_reqstat_LostFocus()
'On Error GoTo l2
''Set D1 = W1.ActiveDocument
''ADDFLAG = False
''REPFLAG = False
''IMPFLAG = False
''STR1 = ""
''STR2 = ""
''STR3 = ""
''
''
''D1.Content = ""
''D1.SaveAs "C:\REPORT.DOC"
''D1.Close
''W1.Quit
'''W1.Visible = False
''Exit Sub
''l2:
End Sub

Private Sub msk_caseno_LostFocus()
Msk_caseno.Text = UCase(Msk_caseno.Text)
If Opt_patient.Value = True Then
    cno_pad = padzero(5, "C", Msk_caseno.Text, 3)
    Msk_caseno.Text = cno_pad
    If SELTA101_FLAG = False Then
        SELTA101_QR(0) = cno_pad
        Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
        SELTA101_FLAG = True
    Else
        SELTA101_QR(0) = cno_pad
        SELTA101_RS.MoreResults
        SELTA101_RS.Requery
    End If
    If SELTA101_RS.RowCount > 0 Then
        txt_name.Visible = True
        lbl_name.Visible = True
        
        txt_name.Text = Trim(SELTA101_RS("TAG")) & " " & Trim(SELTA101_RS("fIRST_NAME")) & " " & Trim(SELTA101_RS("MID_NAME")) & " " & Trim(SELTA101_RS("LAST_NAME"))
    End If
    
    
    mbqrysts_qr(0) = cno_pad
    mbqrysts_qr(1) = LOGIN '"A" + usrid '"A" + logonfrm.txtpassword.Text
    Set mbqrysts_rs = mbqrysts_qr.OpenResultset(1, 3)
    mbqrysts_rs.MoreResults
    mbqrysts_rs.Requery
    
    
    If mbqrysts_rs.RowCount > 1 Then
        cbo_req.clear
        If mbqrysts_rs.EOF = False Then
            cbo_req.AddItem "ALL"
        End If
        
        While mbqrysts_rs.EOF = False
            cbo_req.AddItem req_chglen(mbqrysts_rs(0))
            mbqrysts_rs.MoveNext
        Wend
            
            Exit Sub
    End If

'    If mbqrysts_rs.RowCount > 1 Then
'            cbo_req.clear
'        While mbqrysts_rs.EOF = False
'            cbo_req.AddItem req_chglen(mbqrysts_rs(0))
'            mbqrysts_rs.MoveNext
'        Wend
'            Exit Sub
'    End If
    
    rq_req(0) = cno_pad
    rs_req.MoreResults
    rs_req.Requery
    cbo_req.clear
    If rs_req.EOF = False Then
        cbo_req.AddItem "ALL"
    End If
    While rs_req.EOF = False
        cbo_req.AddItem req_chglen(rs_req(0))
        rs_req.MoveNext
    Wend
Else
    cno_pad = Msk_caseno.FormattedText
    If SELRREQUISITNFLAG1 = False Then
        SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG1 = True
    Else
        SELRREQUISITN_QR(0) = Msk_caseno.FormattedText
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
    End If
    cbo_req.clear
    If SELRREQUISITN_RS.EOF = False Then
        cbo_req.AddItem "ALL"
    End If
    While SELRREQUISITN_RS.EOF = False
        cbo_req.AddItem req_chglen(SELRREQUISITN_RS(0))
        SELRREQUISITN_RS.MoveNext
    Wend
End If

End Sub

Private Sub Opt_all_Click()

   Opt_case.Enabled = False
   Msk_caseno.Enabled = False
   cbo_req.Enabled = False
   cbo_modality.Enabled = True
   
End Sub

Private Sub Opt_case_Click()

   OPT_ALL.Enabled = False
   MaskEd_from_date.Enabled = False
    cbo_modality.Enabled = False
   MaskEd_to_date.Enabled = False
   Msk_caseno.Enabled = True
   cbo_req.Enabled = True
   
    Msk_caseno.SetFocus
   
   ' If Opt_case.Value = True Then
     ' txt_msg.Text = "SELECT CASE NO AND REQUISION NO FROM COMBO BOX"
           
'     End If
'   MaskEd_from_date.SetFocus
'   txt_msg.Text = "ENTER FROM DATE AND TO DATE"
End Sub

Public Function clear()
'    txt_date.Text = ""
'    txt_time.Text = ""
    MaskEd_from_date.Mask = ""
    MaskEd_from_date.Text = ""
    MaskEd_from_date.Mask = "##/##/####"
    
    MaskEd_to_date.Mask = ""
    MaskEd_to_date.Text = ""
    MaskEd_to_date.Mask = "##/##/####"
    Msk_caseno.Text = ""
    cbo_req.Text = ""
    cbo_modality.ListIndex = -1
'    msf_reqstat.clear
    
    OPT_ALL.Enabled = False
    Opt_case.Enabled = False
    Msk_caseno.Enabled = False
    cbo_req.Enabled = False
    msf_reqstat.Rows = 1
    srno = 1
    End Function




Public Sub cmdview()

'     If RIS1_RS("VIEW") = 1 Then
'         CMD_VIEW.Enabled = True
'     Else
'         CMD_VIEW.Enabled = False
'     End If
       
End Sub

Public Sub CMDPRINT()

'     If RIS1_RS("PRINT") = 1 Then
'
'         cmd_print.Enabled = True
'    Else
'         cmd_print.Enabled = False
'
'    End If

End Sub

Private Sub Microbiology(moda As String)
On Error GoTo MSG
Load Report_microbiology
'Report_microbiology.Hide
Report_microbiology.show
If moda = "BA" Then
    Report_microbiology.cmb_modality.Text = "BACTERIOLOGY$BA"
    Report_microbiology.printpage "BA", var_caseno, VAR_REQNNO
ElseIf moda = "MB" Then
    Report_microbiology.cmb_modality.Text = "MYCOBACTERIOLOGY$MB"
    Report_microbiology.printpage "MB", var_caseno, VAR_REQNNO
ElseIf moda = "CM" Then
    Report_microbiology.cmb_modality.Text = "CLINICAL MICROBIOLOGY$CM"
    'Report_microbiology.printpage "CM", var_caseno, VAR_REQNNO
    'CCommented and added by Shraddha on 13/09/2011
    Report_microbiology.printpage "CM", Msk_caseno, VAR_REQNNO
ElseIf moda = "MC" Then
    Report_microbiology.cmb_modality.Text = "MYCOLOGY$MC"
    Report_microbiology.printpage "MC", var_caseno, VAR_REQNNO
ElseIf moda = "MO" Then
    Report_microbiology.cmb_modality.Text = "MOLECULAR MICROBIOLOGY$MO"
    Report_microbiology.printpage "MO", var_caseno, VAR_REQNNO
ElseIf moda = "SE" Then
    Report_microbiology.cmb_modality.Text = "SEROLOGY$SE"
    Report_microbiology.printpage "SE", var_caseno, VAR_REQNNO
ElseIf moda = "XX" Then
    Report_microbiology.cmb_modality.Text = "OTHERS$XX"
    Report_microbiology.printpage "BA", var_caseno, VAR_REQNNO
End If
    'Unload Report_microbiology
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub Echocardiology(moda As String)
On Error GoTo MSG
Load EChocardiologyreport
EChocardiologyreport.Hide
EChocardiologyreport.show
If moda = "2D" Then
    EChocardiologyreport.cmb_modality.Text = "2DECHO$2D"
    EChocardiologyreport.printehco "2D", var_caseno, VAR_REQNNO
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub
Private Sub Endoscopy(moda As String)
On Error GoTo MSG
Load frm_endoscopy1
frm_endoscopy1.Hide
frm_endoscopy1.show
If moda = "ED" Then
    frm_endoscopy1.cmb_modality.Text = "2DECHO$2D"
    frm_endoscopy1.printendo "ED", var_caseno, VAR_REQNNO
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub


Private Sub Opt_patient_Click()
If Opt_patient.Value = True Then
     'lbl_caseno.Caption = "Case No."
     Msk_caseno.Mask = "??>/#####"
    'LBL_DEPENDENT.Visible = False
    'cmb_dependents.Visible = False
     Msk_caseno.Enabled = True
     'msk_caseno.SetFocus
   
End If
End Sub

Private Sub opt_staff_Click()
On Error GoTo MSG:
If opt_staff.Value = True Then
    Opt_case.Caption = "CC NO."
    'lbl_caseno.Caption = "  CC No."
    Msk_caseno.Mask = "########"
    'LBL_DEPENDENT.Visible = True
    'cmb_dependents.Visible = True
    Msk_caseno.Enabled = True
    'msk_caseno.SetFocus
   
'Else
    'lbl_caseno.Caption = "Case No."
End If
'btnenabl ("NEW")
Exit Sub
MSG:
MsgBox Err.Description
End Sub
Private Sub SPREPORT()
If cbo_modality.Text = "ALL" Then
     RQ_ALLSP1(0) = GET_REVDATE(MaskEd_from_date.Text)
     RQ_ALLSP1(1) = GET_REVDATE(MaskEd_to_date.Text)
     Set RS_ALLSP = RQ_ALLSP1.OpenResultset(2, 3)
     cmd_print.Enabled = False
 Else
     RQ_ALLSP(0) = GET_REVDATE(MaskEd_from_date.Text)
     RQ_ALLSP(1) = GET_REVDATE(MaskEd_to_date.Text)
     RQ_ALLSP(2) = Trim(cbo_modality)
     RS_ALLSP.MoreResults
     RS_ALLSP.Requery
     cmd_print.Enabled = False
End If
If RS_ALLSP.RowCount <> 0 Then
    cmd_print.Enabled = True
    RS_ALLSP.MoveFirst
    
    msf_reqstat.Rows = 1
    srno = 1
    While Not RS_ALLSP.EOF
'    While Not rs_grid_det.EOF
      i = RS_ALLSP.RowCount
                        
                        
                        If Trim(RS_ALLSP(1)) = "M" Then
                            appt = "No"
                            Memo = "Yes"
                            payment = "No"
                            EXAM = "No"
                            REPORT = "No"
                        End If
                        
                        If Trim(RS_ALLSP(1)) = "P" Then
                                       'If Trim(rs_grid_det("RECPTNO")) = "NOCHARGE" Or Trim(rs_grid_det("RECPTNO")) = "TRUST" Or Trim(rs_grid_det("RECPTNO")) = "COMPANY" Then
                                        If IsNull(Trim(RS_ALLSP("RECPTNO"))) Then
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "YES"
                                        EXAM = "No"
                                        REPORT = "No"
                                        Else
                                        If Trim(RS_ALLSP("RECPTNO")) = "NOCHARGE" Or Trim(RS_ALLSP("RECPTNO")) = "COMPANY" Or Trim(RS_ALLSP("RECPTNO")) = "TRUST" Or Trim(RS_ALLSP("RECPTNO")) = "ITRUST" Or Trim(RS_ALLSP("RECPTNO")) = "ICOMPANY" Or Trim(RS_ALLSP("RECPTNO")) = "IPATIENT" Or Trim(RS_ALLSP("RECPTNO")) = "VPP" Or Trim(RS_ALLSP("RECPTNO")) = "LR" Or Trim(RS_ALLSP("RECPTNO")) = "BAR" Or Trim(RS_ALLSP("RECPTNO")) = "DrOp" Then
                                        appt = "No"
                                        Memo = "Yes"
                                        payment = "Yes"
                                        EXAM = "No"
                                        REPORT = "No"
                                        Else
                                        If IsNull(Trim(RS_ALLSP("EXAMDT"))) Then
                                        appt = "No"
                                        Memo = "Yes"
                                        payment = "YES"
                                        EXAM = "No"
                                        REPORT = "No"
                                        Else
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "YES"
                                        EXAM = "Yes"
                                        REPORT = "No"
                                        End If
                                        End If
                                        End If
                                   End If
                                    If Trim(RS_ALLSP(1)) = "C" Then
                                        
                                        appt = "Yes"
                                        Memo = "YES"
                                        payment = "YES"
                                        EXAM = "No"
                                        REPORT = "No"
                                        
                                    End If
                                                
                                    If Trim(RS_ALLSP(1)) = "E" Then
                                      If IsNull(Trim(RS_ALLSP("RECPTNO"))) Then
                                        
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yes"
                                        Else
                                        If Trim(RS_ALLSP("RECPTNO")) = "NOCHARGE" Or Trim(RS_ALLSP("RECPTNO")) = "COMPANY" Or Trim(RS_ALLSP("RECPTNO")) = "TRUST" Or Trim(RS_ALLSP("RECPTNO")) = "ITRUST" Or Trim(RS_ALLSP("RECPTNO")) = "ICOMPANY" Or Trim(RS_ALLSP("RECPTNO")) = "IPATIENT" Or Trim(RS_ALLSP("RECPTNO")) = "VPP" Or Trim(RS_ALLSP("RECPTNO")) = "LR" Or Trim(RS_ALLSP("RECPTNO")) = "BAR" Or Trim(RS_ALLSP("RECPTNO")) = "DrOp" Then
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yes"
                                        Else
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "Yes"
                                        EXAM = "Yes"
                                        REPORT = "Yes"
                                        End If
                                        End If
                                   End If
                                    
                                    If Trim(RS_ALLSP(1)) = "F" Then
                                       If IsNull(Trim(RS_ALLSP("RECPTNO"))) Then
                                       
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yesf"
                                       Else
                                        If Trim(RS_ALLSP("RECPTNO")) = "NOCHARGE" Or Trim(RS_ALLSP("RECPTNO")) = "COMPANY" Or Trim(RS_ALLSP("RECPTNO")) = "TRUST" Or Trim(RS_ALLSP("RECPTNO")) = "ITRUST" Or Trim(RS_ALLSP("RECPTNO")) = "ICOMPANY" Or Trim(RS_ALLSP("RECPTNO")) = "IPATIENT" Or Trim(RS_ALLSP("RECPTNO")) = "VPP" Or Trim(RS_ALLSP("RECPTNO")) = "LR" Or Trim(RS_ALLSP("RECPTNO")) = "BAR" Or Trim(RS_ALLSP("RECPTNO")) = "DrOp" Then
                                         appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yesf"
                                        Else
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "YES"
                                        EXAM = "Yes"
                                        REPORT = "Yesf"
                                        End If
                                       End If
                                    End If
                                    If Trim(RS_ALLSP(1)) = "T" Then
                                      If IsNull(Trim(RS_ALLSP("RECPTNO"))) Then
                                        
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yest"
                                      Else
                                      If Trim(RS_ALLSP("RECPTNO")) = "NOCHARGE" Or Trim(RS_ALLSP("RECPTNO")) = "COMPANY" Or Trim(RS_ALLSP("RECPTNO")) = "TRUST" Or Trim(RS_ALLSP("RECPTNO")) = "ITRUST" Or Trim(RS_ALLSP("RECPTNO")) = "ICOMPANY" Or Trim(RS_ALLSP("RECPTNO")) = "IPATIENT" Or Trim(RS_ALLSP("RECPTNO")) = "VPP" Or Trim(RS_ALLSP("RECPTNO")) = "LR" Or Trim(RS_ALLSP("RECPTNO")) = "BAR" Or Trim(RS_ALLSP("RECPTNO")) = "DrOp" Then
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yest"
                                        Else
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "YES"
                                        EXAM = "Yes"
                                        REPORT = "Yest"
                                      End If
                                    End If
                                    End If
                        If IsNull(Trim(RS_ALLSP(1))) Then
                            appt = "No"
                            Memo = "No"
                            payment = "No"
                            EXAM = "No"
                            REPORT = "No"
                       End If


'         msf_reqstat.AddItem srno & vbTab & cbo_req.Text & vbTab & rs_grid_det(1) & vbTab & Format(rs_grid_det(2), "dd/mm/yyyy") & vbTab & rs_grid_det(3) & vbTab & rs_grid_det(4)
'            msf_reqstat.AddItem srno & vbTab & cbo_req.Text & vbTab & rs_grid_det(1) & vbTab & appt & vbTab & Memo & vbTab & payment & vbTab & exam & vbTab & report
'            rs_grid_det.MoveNext
'            srno = srno + 1
'                    If IsNull(Len(Trim(rs_allreq(2)))) Then
'                            appt = "no"
'                        Else
'                            appt = "yes"
'                        End If
                               'msf_reqstat.AddItem srno & vbTab & rs_allreq(0) & vbTab & rs_allreq(1) & vbTab & Format(rs_allreq(2), "dd/mm/yyyy") & vbTab & rs_allreq(3) & vbTab & rs_allreq(4)
'                               msf_reqstat.AddItem srno & vbTab & rs_allreq(0) & vbTab & rs_allreq(1) & vbTab & appt & vbTab & rs_allreq(3) & vbTab & rs_allreq(4)

                             'msf_reqstat.AddItem srno & vbTab & req_chglen(rs_allreq(0)) & vbTab & rs_allreq("CASENO") & vbTab & appt & vbTab & Memo & vbTab & payment & vbTab & exam & vbTab & Report
                             msf_reqstat.Rows = msf_reqstat.Rows + 1
                             msf_reqstat.row = srno
                             msf_reqstat.col = 0
                             msf_reqstat.TextMatrix(srno, 0) = srno
                             msf_reqstat.TextMatrix(srno, 1) = Trim(req_chglen(RS_ALLSP(0))) & "  " & Trim(RS_ALLSP("reqndt"))
                             msf_reqstat.TextMatrix(srno, 2) = RS_ALLSP("caseno")
                             
                             
                             msf_reqstat.col = 3
                             If UCase(Memo) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                             
                             msf_reqstat.TextMatrix(srno, 3) = RS_ALLSP("MEMONO") & " " & RS_ALLSP("MEMODT") & ""
                             End If
                            msf_reqstat.col = 4
                             If UCase(payment) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                             msf_reqstat.TextMatrix(srno, 4) = RS_ALLSP("RECPTDT") & ""
                             Else
                              msf_reqstat.CellBackColor = vbGreen
                               If Trim(RS_ALLSP("Recptno")) = "NOCHARGE" Then
                                msf_reqstat.TextMatrix(srno, 4) = "NOCHARGE"
                               End If

                               If Trim(RS_ALLSP("RECPTNO")) = "TRUST" Then
                                msf_reqstat.TextMatrix(srno, 4) = "TRUST"
                               End If
                               
                               If Trim(RS_ALLSP("RECPTNO")) = "COMPANY" Then
                                msf_reqstat.TextMatrix(srno, 4) = "COMPANY"
                               End If
                                
                               If Trim(RS_ALLSP("RECPTNO")) = "ICOMPANY" Then
                                msf_reqstat.TextMatrix(srno, 4) = "ICOMPANY"
                               End If
                               If Trim(RS_ALLSP("RECPTNO")) = "ITRUST" Then
                                msf_reqstat.TextMatrix(srno, 4) = "ITRUST"
                               End If
                               If Trim(RS_ALLSP("RECPTNO")) = "IPATIENT" Then
                                msf_reqstat.TextMatrix(srno, 4) = "IPATIENT"
                               End If
                               If Trim(RS_ALLSP("RECPTNO")) = "VPP" Then
                                msf_reqstat.TextMatrix(srno, 4) = "VPP"
                               End If
                               If Trim(RS_ALLSP("RECPTNO")) = "LR" Then
                                msf_reqstat.TextMatrix(srno, 4) = "LR"
                               End If
                               If Trim(RS_ALLSP("RECPTNO")) = "BAR" Then
                                msf_reqstat.TextMatrix(srno, 4) = "BAR"
                               End If
                               If Trim(RS_ALLSP("RECPTNO")) = "DrOp" Then
                                msf_reqstat.TextMatrix(srno, 4) = "DrOp"
                               End If
                               
                               If IsNull(Trim(RS_ALLSP("RECPTNO"))) Then
                                msf_reqstat.CellBackColor = vbWhite
                                msf_reqstat.TextMatrix(srno, 4) = " "
                               End If
                             
                                                      
                             
                             End If
                             msf_reqstat.col = 5
                             If UCase(appt) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                             msf_reqstat.TextMatrix(srno, 5) = RS_ALLSP("APNTDT") & ""
                             End If
                             
                             msf_reqstat.col = 6
                             If UCase(EXAM) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                             msf_reqstat.TextMatrix(srno, 6) = RS_ALLSP("EXAMDT") & ""
                             End If
                             
                             
                             msf_reqstat.col = 7
                             If UCase(REPORT) = "YESF" Then
                            msf_reqstat.CellBackColor = vbGreen
                            msf_reqstat.TextMatrix(srno, 7) = "FNL " & RS_ALLSP("FREPORTDT")
                             End If
                            If UCase(REPORT) = "YEST" Then
                            msf_reqstat.CellBackColor = vbGreen
                            msf_reqstat.TextMatrix(srno, 7) = "PRV " & RS_ALLSP("PREPORTDT")
                             End If
                             
                              msf_reqstat.col = 8
                            If Mid(RS_ALLSP("REQNNO"), 4, 2) = "IH" Then
                                
                                seldrequisitn_qr(0) = RS_ALLSP("REQNNO")
                                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                                If seldrequisitn_rs.RowCount > 0 Then
                                    If Len(Trim(seldrequisitn_rs("SPIHREQNO"))) > 0 Then
                                    
                                         msf_reqstat.CellBackColor = vbGreen
                                         msf_reqstat.TextMatrix(srno, 8) = req_chglen(Trim(seldrequisitn_rs("spihreqno")))
                                     End If
                                End If
                            End If
                            If Mid(RS_ALLSP("reqnno"), 4, 2) = "SP" Then
                                
                                'SELDREQNOQR1(0) = RS_ALLSP("REQNNO")
                                'Set SELDREQNORS1 = SELDREQNOQR1.OpenResultset(1, 3)
                                'If SELDREQNORS1.RowCount > 0 Then
                                '    msf_reqstat.CellBackColor = vbGreen
                                '    msf_reqstat.TextMatrix(srno, 8) = req_chglen(SELDREQNORS1("REQNO"))
                                'End If
                           '''
                                                '''
                             ' THE CODE SLICE IS COMMENTED BECAUSE THE REQNO IS NOT REQUIRED BY PENDING REMARKS
      ' FIELD FOR THE SP IS REQUIRED FROM SITEDETAIL TABLE OF DISLIB
                                SELSITEDETRQR1(0) = RS_ALLSP("reqnno")
                           
                                Set SELSITEDETRRS = SELSITEDETRQR1.OpenResultset(1, 3)
                                If SELSITEDETRRS.RowCount > 0 Then
                                    msf_reqstat.CellBackColor = vbGreen
                                    msf_reqstat.TextMatrix(srno, 8) = SELSITEDETRRS("FLD1")
                                End If
                           '''
                            End If
                            msf_reqstat.col = 9
                            seldrequisitn_qr(0) = (RS_ALLSP("REQNNO"))
                            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                            If seldrequisitn_rs.RowCount > 0 Then
                                msf_reqstat.TextMatrix(srno, 9) = Trim(seldrequisitn_rs("PATHNO"))
                            End If
                             
                                RS_ALLSP.MoveNext
                               srno = srno + 1
                    Wend
                '    MsgBox msf_reqstat.Rows
                Else
                
                MsgBox "NO DATA is found between these two dates."
                  button_ena_dis ("cancel")
                  clear
                 MaskEd_from_date.Enabled = False
                 MaskEd_to_date.Enabled = False
                 cmd_view.SetFocus
            End If
End Sub
Private Sub HLREPORT()

selrrequisitnqr(0) = VAR_REQNNO
Set selrrequisitnrs = selrrequisitnqr.OpenResultset(1, 3)

If SELDCBCREPORT_FLAG = False Then
    SELDCBCREPORT_QR(0) = VAR_REQNNO
    Set SELDCBCREPORT_RS = SELDCBCREPORT_QR.OpenResultset(1, 3)
    SELDCBCREPORT_FLAG = True
Else
    SELDCBCREPORT_QR(0) = VAR_REQNNO
    SELDCBCREPORT_RS.MoreResults
    SELDCBCREPORT_RS.Requery
End If
'If SELDCBCREPORT_RS.RowCount = 0 Then
'  Exit Sub
'End If

For i = 0 To 25
    CrystalReport1.Formulas(i) = ""
Next i
selta1012_qr(0) = selrrequisitnrs("CASENO")
Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)

If Opt_patient.Value = True Then
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
    'If Trim(txt_modcode.Text) = "HL" Then
    CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShl.RPT"
    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(Left(msf_reqstat.TextMatrix(msf_reqstat.row, 1), 16)) & "'"
'       If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'       Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'       End If
    CrystalReport1.Formulas(2) = "PROVDIA = '" & selta1012_rs("PRVDIAGDES") & "'"
    If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
        CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
        str1 = ""
        str1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
        If Len(Trim(selrrequisitnrs("enterby"))) > 0 Then
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
        End If
    Else
        CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
        CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
    End If
    CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
             If Trim(SELDCBCREPORT_RS("lap")) = "" And Trim(SELDCBCREPORT_RS("MPO")) = "" And Trim(SELDCBCREPORT_RS("SBB")) = "" And Trim(SELDCBCREPORT_RS("PSA")) = "" And Trim(SELDCBCREPORT_RS("NSE")) = "" And Trim(SELDCBCREPORT_RS("TRAD")) = "" Then
                'If CYTOCHEMFLAG = True Then
                CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
            Else
                CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
                'End If
            End If
            If Trim(SELDCBCREPORT_RS("BT")) = "" And Trim(SELDCBCREPORT_RS("PT")) = "" And Trim(SELDCBCREPORT_RS("PLATELET1")) = "" And Trim(SELDCBCREPORT_RS("PTTK")) = "" And Trim(SELDCBCREPORT_RS("TT")) = "" And Trim(SELDCBCREPORT_RS("FIBRINOGEN")) = "" And Trim(SELDCBCREPORT_RS("FDPDDIMER")) = "" And Trim(SELDCBCREPORT_RS("OTHERS")) = "" Then
                CrystalReport1.Formulas(11) = "CP = '" & "" & "'"
            Else
                CrystalReport1.Formulas(11) = "CP = '" & "ABC" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("NORMBOBLAS")) = "" And Trim(SELDCBCREPORT_RS("OTHERCELLR")) = "" And Trim(SELDCBCREPORT_RS("rbcmorpho")) = "" And Trim(SELDCBCREPORT_RS("reticcount")) = "" And Trim(SELDCBCREPORT_RS("mparasite")) = "" And Trim(SELDCBCREPORT_RS("ESR")) = "" Then
                CrystalReport1.Formulas(15) = "RBC = '" & "" & "'"
            Else
                CrystalReport1.Formulas(15) = "RBC = '" & "ABC" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("PMORPHOLOG")) = "" And Trim(SELDCBCREPORT_RS("FLD1")) = "" Then
                CrystalReport1.Formulas(16) = "PLATELETS = '" & "" & "'"
            Else
                CrystalReport1.Formulas(16) = "PLATELETS = '" & "ABC" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("PROLYMPHOC")) = "" And Trim(SELDCBCREPORT_RS("Sezarycell")) = "" And Trim(SELDCBCREPORT_RS("hairycells")) = "" And Trim(SELDCBCREPORT_RS("othercells")) = "" Then
                CrystalReport1.Formulas(17) = "prolyn = '" & "" & "'"
            Else
                CrystalReport1.Formulas(17) = "prolyn = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("blasts")) = "" And Trim(SELDCBCREPORT_RS("PROMYELOCY")) = "" And Trim(SELDCBCREPORT_RS("myelocytes")) = "" And Trim(SELDCBCREPORT_RS("metamyeloc")) = "" And Trim(SELDCBCREPORT_RS("plasmacell")) = "" And Trim(SELDCBCREPORT_RS("aTlymphEE")) = "" Then
                CrystalReport1.Formulas(18) = "blast = '" & "" & "'"
            Else
                CrystalReport1.Formulas(18) = "blast = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("neutrophi1")) = "" And Trim(SELDCBCREPORT_RS("lymphocyt1")) = "" And Trim(SELDCBCREPORT_RS("monocytE1")) = "" And Trim(SELDCBCREPORT_RS("eosinophi1")) = "" And Trim(SELDCBCREPORT_RS("basophill1")) = "" Then
                CrystalReport1.Formulas(19) = "cbcm = '" & "" & "'"
            Else
                CrystalReport1.Formulas(19) = "cbcm = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("HCT")) = "" And Trim(SELDCBCREPORT_RS("MCH")) = "" And Trim(SELDCBCREPORT_RS("MCV")) = "" And Trim(SELDCBCREPORT_RS("MCHC")) = "" And Trim(SELDCBCREPORT_RS("MPV")) = "" And Trim(SELDCBCREPORT_RS("RDW")) = "" And Trim(SELDCBCREPORT_RS("ABNORMAL")) = "" Then
                CrystalReport1.Formulas(20) = "rbci = '" & "" & "'"
            Else
                CrystalReport1.Formulas(20) = "rbci = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("neutrophil")) = "" And Trim(SELDCBCREPORT_RS("lymphocyte")) = "" And Trim(SELDCBCREPORT_RS("monocytes")) = "" And Trim(SELDCBCREPORT_RS("eosinophil")) = "" And Trim(SELDCBCREPORT_RS("basophils")) = "" And Trim(SELDCBCREPORT_RS("immgranulo")) = "" Then
                CrystalReport1.Formulas(21) = "dlc = '" & "" & "'"
            Else
                CrystalReport1.Formulas(21) = "dlc = '" & "123" & "'"
            End If
            If Trim(SELDCBCREPORT_RS("HB")) = "" And Trim(SELDCBCREPORT_RS("RBC")) = "" And Trim(SELDCBCREPORT_RS("TLC")) = "" And Trim(SELDCBCREPORT_RS("platelets")) = "" Then
                CrystalReport1.Formulas(22) = "cbcc = '" & "" & "'"
            Else
                CrystalReport1.Formulas(22) = "cbcc = '" & "123" & "'"
            End If
End If
If opt_staff.Value = True Then
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
    If selrrequisitnrs("MODALITYCD") = "EK" Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISECS.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'           If Trim(prepath) = "" Then
'           CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'           Else
'           CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
    End If
End If
'***** STAFF
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport1.Action = 1

End Sub

Private Sub EKREPORT()
selrrequisitnqr(0) = VAR_REQNNO
Set selrrequisitnrs = selrrequisitnqr.OpenResultset(1, 3)

If selrrequisitnrs("MODALITYCD") = "EK" Then
    If Opt_patient.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEC.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        'CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "ECG REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
            If Len(Trim(selrrequisitnrs("enterby"))) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                 CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
            End If
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "ECG REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        If SELECREPORTFLAG = False Then
            SELECREPORTQR(0) = VAR_REQNNO
            Set SELECREPORTRS = SELECREPORTQR.OpenResultset(1, 3)
            SELECREPORTFLAG = True
        Else
            SELECREPORTQR(0) = VAR_REQNNO
            SELECREPORTRS.MoreResults
            SELECREPORTRS.Requery
        End If
        If Len(Trim(SELECREPORTRS("RHYTHM"))) > 3 Then
            CrystalReport1.Formulas(13) = "rhythm = '" & "Rhythm" & "'"
        Else
            CrystalReport1.Formulas(13) = "rhythm = '" & "" & "'"
        End If
        If (Trim(SELECREPORTRS("PWAVE"))) <> "" Then
            CrystalReport1.Formulas(14) = "pwave = '" & "Pwave" & "'"
        Else
            CrystalReport1.Formulas(14) = "pwave = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("TWAVE"))) > 3 Then
            CrystalReport1.Formulas(15) = "twave = '" & "Twave" & "'"
        Else
            CrystalReport1.Formulas(15) = "twave = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("STSEGMENT"))) > 3 Then
            CrystalReport1.Formulas(16) = "stsegment = '" & "ST Segment" & "'"
        Else
            CrystalReport1.Formulas(16) = "stsegment = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("RATE"))) > 0 Then
            CrystalReport1.Formulas(17) = "rate = '" & "Rate" & "'"
        Else
            CrystalReport1.Formulas(17) = "rate = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("ATRIALRATE"))) > 0 Then
            CrystalReport1.Formulas(18) = "arate = '" & "Atrial Rate" & "'"
        Else
            CrystalReport1.Formulas(18) = "arate = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("VENRATE"))) > 0 Then
            CrystalReport1.Formulas(19) = "vrate = '" & "Ventricular Rate" & "'"
        Else
            CrystalReport1.Formulas(19) = "vrate = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("PRINTERVAL"))) > 0 Then
            CrystalReport1.Formulas(20) = "printerval = '" & "PR Interval" & "'"
        Else
            CrystalReport1.Formulas(20) = "printerval = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("QRSMORPH"))) > 0 Then
            CrystalReport1.Formulas(21) = "qrsm = '" & "QRS Morphology" & "'"
        Else
            CrystalReport1.Formulas(21) = "qrsm = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("QRSDUR"))) > 0 Then
            CrystalReport1.Formulas(22) = "qrd = '" & "QRS Duration" & "'"
        Else
            CrystalReport1.Formulas(22) = "qrd = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("QTINTERVAL"))) > 0 Then
            CrystalReport1.Formulas(23) = "qti = '" & "QT Interval" & "'"
        Else
            CrystalReport1.Formulas(23) = "qti = '" & "" & "'"
        End If
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEC.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        'CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "ECG REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
            If Len(Trim(selrrequisitnrs("enterby"))) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                 CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
            End If
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "ECG REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        If SELECREPORTFLAG = False Then
            SELECREPORTQR(0) = VAR_REQNNO
            Set SELECREPORTRS = SELECREPORTQR.OpenResultset(1, 3)
            SELECREPORTFLAG = True
        Else
            SELECREPORTQR(0) = VAR_REQNNO
            SELECREPORTRS.MoreResults
            SELECREPORTRS.Requery
        End If
        If Len(Trim(SELECREPORTRS("RHYTHM"))) > 3 Then
            CrystalReport1.Formulas(13) = "rhythm = '" & "Rhythm" & "'"
        Else
            CrystalReport1.Formulas(13) = "rhythm = '" & "" & "'"
        End If
        If (Trim(SELECREPORTRS("PWAVE"))) <> "" Then
            CrystalReport1.Formulas(14) = "pwave = '" & "Pwave" & "'"
        Else
            CrystalReport1.Formulas(14) = "pwave = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("TWAVE"))) > 3 Then
            CrystalReport1.Formulas(15) = "twave = '" & "Twave" & "'"
        Else
            CrystalReport1.Formulas(15) = "twave = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("STSEGMENT"))) > 3 Then
            CrystalReport1.Formulas(16) = "stsegment = '" & "ST Segment" & "'"
        Else
            CrystalReport1.Formulas(16) = "stsegment = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("RATE"))) > 0 Then
            CrystalReport1.Formulas(17) = "rate = '" & "Rate" & "'"
        Else
            CrystalReport1.Formulas(17) = "rate = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("ATRIALRATE"))) > 0 Then
            CrystalReport1.Formulas(18) = "arate = '" & "Atrial Rate" & "'"
        Else
            CrystalReport1.Formulas(18) = "arate = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("VENRATE"))) > 0 Then
            CrystalReport1.Formulas(19) = "vrate = '" & "Ventricular Rate" & "'"
        Else
            CrystalReport1.Formulas(19) = "vrate = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("PRINTERVAL"))) > 0 Then
            CrystalReport1.Formulas(20) = "printerval = '" & "PR Interval" & "'"
        Else
            CrystalReport1.Formulas(20) = "printerval = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("QRSMORPH"))) > 0 Then
            CrystalReport1.Formulas(21) = "qrsm = '" & "QRS Morphology" & "'"
        Else
            CrystalReport1.Formulas(21) = "qrsm = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("QRSDUR"))) > 0 Then
            CrystalReport1.Formulas(22) = "qrd = '" & "QRS Duration" & "'"
        Else
            CrystalReport1.Formulas(22) = "qrd = '" & "" & "'"
        End If
        If Len(Trim(SELECREPORTRS("QTINTERVAL"))) > 0 Then
            CrystalReport1.Formulas(23) = "qti = '" & "QT Interval" & "'"
        Else
            CrystalReport1.Formulas(23) = "qti = '" & "" & "'"
        End If
        CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
    End If
    CrystalReport1.WindowState = crptMaximized
'    CRYSTALREPORT1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
End If

End Sub
Private Sub BFREPORT()
For i = 0 To 25
    CrystalReport1.Formulas(i) = ""
Next i
VAR_REQNNO = req_chglen(Left(msf_reqstat.Text, 16))
selta1012_qr(0) = selrrequisitnrs("CASENO")
Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)

selrrequisitnqr(0) = VAR_REQNNO
Set selrrequisitnrs = selrrequisitnqr.OpenResultset(1, 3)
If Opt_patient.Value = True Then
             
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"

    
    
     If Trim(selrrequisitnrs("MODALITYCD")) = "BF" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(selrrequisitnrs("ENTERBY"))
            If Len(Trim(selrrequisitnrs("ENTERBY"))) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                 CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
            End If
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        
        
     End If
       
End If
If opt_staff.Value = True Then
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"

 
     If Trim(selrrequisitnrs("MODALITYCD")) = "BF" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBFS.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & VAR_REQNNO & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "BODY FLUID REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        
        
     End If

End If
'***** STAFF
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CrystalReport1.Action = 1
End Sub

Public Sub hlreport1()
'On Error GoTo MSG:

Dim str1 As String
Dim str2 As String
Dim STR3 As String
Dim REP1 As Boolean
Dim REP2 As Boolean

Dim i As Integer
' On Error GoTo prtmsg:

DoEvents

'Dim VAR_REQNNO As String
Dim var_caseno As String

REP1 = False
REP2 = False
VAR_REQNNO = req_chglen(Left(msf_reqstat.Text, 16))  '''req_chglen(mask_edit_req.Text)
var_caseno = msf_reqstat.TextMatrix(msf_reqstat.row, 2)
'VAR_REQNNO = reqno
SELDHEMCYTO_QR(0) = VAR_REQNNO
Set SELDHEMCYTO_RS = SELDHEMCYTO_QR.OpenResultset(1, 1)
If SELDHEMCYTO_RS.RowCount > 0 Then

 If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("HEM")) = "YES" Or Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
    REP1 = True
    
 End If
  If Trim(SELDHEMCYTO_RS("CYTO")) = "YES" Or Trim(SELDHEMCYTO_RS("SMAR")) = "YES" Then
    REP2 = True
  End If
End If
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
       ' VAR_REQNNO = req_chglen(cmb_reqno.Text)
        If REFDRFLAG1 = False Then
                REFDR_QR(0) = var_caseno ' Trim(MaskEdBox1.FormattedText)
                REFDR_QR(1) = VAR_REQNNO  ' Trim(req_chglen(cmb_reqno.Text))
                Set REFDR_RS = REFDR_QR.OpenResultset(1, 3)
                REFDRFLAG1 = True
            Else
                REFDR_QR(0) = var_caseno   'Trim(MaskEdBox1.FormattedText)
                REFDR_QR(1) = VAR_REQNNO  ' Trim(req_chglen(cmb_reqno.Text))
                REFDR_RS.MoreResults
                REFDR_RS.Requery
            End If


'
If Opt_patient.Value = True Then
    If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
    
        '***** Shraddha 10/12/10 *****
        'str = "SELECT * FROM DISLIB.HINTERFACE WHERE REQNO=?"
        SEL_DHINTER_QR(0) = VAR_REQNNO
        Set SEL_DHINTER_RS = SEL_DHINTER_QR.OpenResultset(1, 3)
        If Not SEL_DHINTER_RS.EOF Then
        '*****************************
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
        
        'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' "
        CrystalReport1.SelectionFormula = "{reporthlf1.REQNO} = '" & VAR_REQNNO & "' "
        'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISAUTOCBC.RPT"
        'Commented and added by Shraddha on 13/09/10
        CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISAUTOCBC.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "F" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "AUTOMATED COMPLETE BLOOD COUNT REPORT" & "'"
            str1 = ""
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "AUTOMATED COMPLETE BLOOD COUNT REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        
        '***** Shraddha 04/01/11 *****
                    
            sel_dhemcbc_qr(0) = Trim(VAR_REQNNO)
            sel_dhemcbc_qr(1) = "R"
            Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 1)
            
            Dim strRBC, strWBC, strPLATELET As String
            strRBC = ""
            strWBC = ""
            strPLATELET = ""
            While Not sel_dhemcbc_rs.EOF
                'strRBC = strRBC & Trim(SEL_DHEMCBC_RS("PRO_CODE")) & ":" & Trim(SEL_DHEMCBC_RS("OBSERVALUE")) & " "
                strRBC = strRBC & Trim(sel_dhemcbc_rs("TESTNAME")) & ":" & Trim(sel_dhemcbc_rs("OBSERVALUE")) & " "
                sel_dhemcbc_rs.MoveNext
            Wend
            
            CrystalReport1.Formulas(13) = "RBC = '" & strRBC & "'"
            
            '*****
            sel_dhemcbc_qr(0) = Trim(VAR_REQNNO)
            sel_dhemcbc_qr(1) = "W"
            Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 1)
                        
            While Not sel_dhemcbc_rs.EOF
                'strWBC = strWBC & Trim(SEL_DHEMCBC_RS("PRO_CODE")) & ":" & Trim(SEL_DHEMCBC_RS("OBSERVALUE")) & " "
                strWBC = strWBC & Trim(sel_dhemcbc_rs("TESTNAME")) & ":" & Trim(sel_dhemcbc_rs("OBSERVALUE")) & " "
                sel_dhemcbc_rs.MoveNext
            Wend
            
            CrystalReport1.Formulas(14) = "WBC = '" & strWBC & "'"
            
            '*****
            sel_dhemcbc_qr(0) = Trim(VAR_REQNNO)
            sel_dhemcbc_qr(1) = "P"
            Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 1)
                        
            While Not sel_dhemcbc_rs.EOF
                'strPLATELET = strPLATELET & Trim(SEL_DHEMCBC_RS("PRO_CODE")) & ":" & Trim(SEL_DHEMCBC_RS("OBSERVALUE")) & " "
                strPLATELET = strPLATELET & Trim(sel_dhemcbc_rs("TESTNAME")) & ":" & Trim(sel_dhemcbc_rs("OBSERVALUE")) & " "
                sel_dhemcbc_rs.MoveNext
            Wend
            
            CrystalReport1.Formulas(15) = "PLATELET = '" & strPLATELET & "'"
        
        '*****************************
        
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
        
        '***** Shraddha 15/04/11 *****
        SEL_DHEMCBCDET_QR(0) = Trim(VAR_REQNNO)
        Set SEL_DHEMCBCDET_RS = SEL_DHEMCBCDET_QR.OpenResultset(1, 1)
        If Not SEL_DHEMCBCDET_RS.EOF Then
            While Not SEL_DHEMCBCDET_RS.EOF
                    If Trim(SEL_DHEMCBCDET_RS("PRO_CODE")) = "RETIC %" Or Trim(SEL_DHEMCBCDET_RS("PRO_CODE")) = "RETIC ABS" Or Trim(SEL_DHEMCBCDET_RS("PRO_CODE")) = "MCVR" Or Trim(SEL_DHEMCBCDET_RS("PRO_CODE")) = "CHCMR" Or Trim(SEL_DHEMCBCDET_RS("PRO_CODE")) = "CHR" Then
                        ret = True
                    End If
                    SEL_DHEMCBCDET_RS.MoveNext
            Wend
            If ret = True Then
                str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                CrystalReport6.SelectionFormula = "{REPORTHLF1.REQNO} = '" & VAR_REQNNO & "'"
                CrystalReport6.ReportFileName = "C:\DisreportsecLF\DISAUTOCBCretic.RPT"
                'CrystalReport6.Formulas(2) = "PROVDIA = '" & selta1012_rs("PRVDIAGDES") & "'"
                CrystalReport6.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
                CrystalReport6.Formulas(0) = "REQNO = '" & Trim(VAR_REQNNO) & "'"
                CrystalReport6.Formulas(5) = "PROFIN = '" & "AUTOMATED COMPLETE BLOOD COUNT REPORT" & "'"
                CrystalReport6.Formulas(12) = "USRID = '" & str1 & "'"
                CrystalReport6.Destination = crptToWindow
                'CrystalReport3.Destination = crptToPrinter
                CrystalReport6.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport6.Action = 1
            End If
        End If
        '**************** 15/04/11 *****************
        
        
        Else
        '***** Shraddha 10/12/10 *****
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
        
        CrystalReport1.SelectionFormula = "{reporthllf.REQNO} = '" & VAR_REQNNO & "' "
        'CRYSTALREPORT1.SelectionFormula = "{reporthlf1.REQNO} = '" & VAR_REQNNO & "' "
        CrystalReport1.ReportFileName = "C:\DISREPORTSEClf\lh report\DISAUTOCBC.RPT"
        'Commented and added by Shraddha on 13/09/10
        'CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "F" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "AUTOMATED COMPLETE BLOOD COUNT REPORT" & "'"
            str1 = ""
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "AUTOMATED COMPLETE BLOOD COUNT REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
        '*****************************
        End If
        
    End If  'CBC/HEM/CG
    
    
    '############## Shraddha 09/10/09 ###############
    
    ''** FOR COAGULATION PROFILE

    SEL_DHEMCPDET_QR(0) = VAR_REQNNO
    Set SEL_DHEMCPDET_RS = SEL_DHEMCPDET_QR.OpenResultset(1, 1)
    
    If Not SEL_DHEMCPDET_RS.EOF Then
    
    Dim a As String
        
    

        CrystalReport1.Formulas(0) = "REQNO='" & Trim(VAR_REQNNO) & "'"
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    
        CrystalReport1.Formulas(1) = "USRID='" & str & "'"
        CrystalReport1.Formulas(2) = "PROFIN='" & "HAEMATO PATHOLOGY REPORT" & "'"
        
        CrystalReport1.Formulas(3) = "nocomments='" & "YES" & "'"
        CrystalReport1.Formulas(4) = "COMMENTS='" & "" & "'"
        
        CrystalReport1.Formulas(5) = "noimpress='" & "YES" & "'"
        CrystalReport1.Formulas(6) = "impress='" & "" & "'"
        
        str = ""
        SELCPDETVIEW_QR(0) = Trim(VAR_REQNNO)
        Set SELCPDETVIEW_RS = SELCPDETVIEW_QR.OpenResultset(2, 1)
        
        While Not SELCPDETVIEW_RS.EOF
        
            If Trim(SELCPDETVIEW_RS("PROCCODE")) = "APTT" Then
                If Trim(SELCPDETVIEW_RS("OBSERVALUE")) = "" Then
                    a = "blank"
                Else
                    a = "filled"
                End If
            End If
            
            If Trim(SELCPDETVIEW_RS("PROCCODE")) = "D-DIMER" Then
                CrystalReport1.Formulas(7) = "D-DIMER='" & "D-Dimer test is performed manually by Semi-Quantitative Latex Agglutination Technique." & "'"
            End If
            
            If Trim(SELCPDETVIEW_RS("PROCCODE")) = "TT" Then
                CrystalReport1.Formulas(8) = "TT='" & "TT test is performed manually." & "'"
            End If
            
            If Trim(SELCPDETVIEW_RS("PROCCODE")) = "PT" Then
                str = str & "/" & "PT"
            ElseIf Trim(SELCPDETVIEW_RS("PROCCODE")) = "APTT" Then
                str = str & "/" & "APTT"
            ElseIf Trim(SELCPDETVIEW_RS("PROCCODE")) = "FIBRINOGEN" Then
                str = str & "/" & "FIBRINOGEN"
            ElseIf Trim(SELCPDETVIEW_RS("PROCCODE")) = "INR" Then
                str = str & "/" & "INR"
            End If
        
            str1 = "ACL01"
        
            CrystalReport1.Formulas(9) = "PT='" & str & " tests are performed on automated coagulometer " & str1 & ".'"
        
            SELCPDETVIEW_RS.MoveNext
        Wend
        
        
        If Len(Trim(msf_reqstat.TextMatrix(msf_reqstat.row, 2))) <> 6 Then 'Patient
            CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISCPDET_INTER.rpt"
            
            selrrequisitnqr(0) = Trim(VAR_REQNNO)
            Set selrrequisitnrs = selrrequisitnqr.OpenResultset(1, 3)
            If Not selrrequisitnrs.EOF Then
                CrystalReport1.Formulas(10) = "signature='" & Trim(selrrequisitnrs("freportby")) & "'"
            End If
        Else ''Staff report
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCPDETSTAFF_INTER.rpt"
        End If
    
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO}='" & Trim(VAR_REQNNO) & "'"
        CrystalReport1.Destination = crptToWindow
        'CRYSTALREPORT1.Destination = crptToPrinter
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB,PATOOLSLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        
        If a = "filled" Then
        
        CrystalReport1.Action = 1
        
        End If
    
    End If
    
    '################################################
    
    ''***** Added by Shraddha on 11/03/11 *****
    'if dcbcrep1.esr has value then
    
    SEL_DCBCREP1_QR(0) = Trim(VAR_REQNNO)
    Set SEL_DCBCREP1_RS = SEL_DCBCREP1_QR.OpenResultset(2, 1)
        
    If Not SEL_DCBCREP1_RS.EOF Then
    
    If Trim(SEL_DCBCREP1_RS("ESR")) <> "" Then
        CrystalReport5.SelectionFormula = "{RREQUISITN.REQNNO} = '" & Trim(VAR_REQNNO) & "'"
        CrystalReport5.ReportFileName = "c:\DISREPORTSECLF\DISAUTOCBC_ESR.rpt"
        
        CrystalReport5.Formulas(0) = "REQNO='" & Trim(VAR_REQNNO) & "'"
        CrystalReport5.Formulas(1) = "PROVDIA = '" & txt_procedure & "'"
        CrystalReport5.Formulas(2) = "NAM = '" & txt_name & "'"
        
        If SELTA101_FLAG = False Then
            SELTA101_QR(0) = selrrequisitnrs("CASENO")
            Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
        End If
            
        If Not SELTA101_RS.EOF Then
            'txt_sex.Text = SELTA101_RS("sex")
            'txt_age = DateDiff("m", SELTA101_RS("BIRTH_DATE"), CStr(DATE11))
            CrystalReport5.Formulas(3) = "Age = '" & Trim(SELTA101_RS("sex")) & "'"
            CrystalReport5.Formulas(4) = "Sex = '" & DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), CStr(DATE11)) & "'"
        End If
        
        CrystalReport5.Formulas(5) = "USRID = '" & UI & "'"
        
        CrystalReport5.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport5.WindowState = crptMaximized
        CrystalReport5.Action = 1
        
    End If
    End If
    '******************************************
    
    
    '''by Kanchan
    
    'If Trim(SELDHEMCYTO_RS("HEM")) = "YES" Then
        
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
        
        CrystalReport4.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' "
        
        CrystalReport4.ReportFileName = "C:\DISREPORTSEC\HEMATO.RPT" 'PBS
        
        CrystalReport4.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        'CrystalReport1.Formulas(1) = "PROVDIA = '" & txt_procedure & "'"
            
        '***** Shraddha 09/05/11 *****
          selrrequisitnqr(0) = Trim(VAR_REQNNO)
            Set selrrequisitnrs = selrrequisitnqr.OpenResultset(1, 3)
            If Not selrrequisitnrs.EOF Then
                CrystalReport4.Formulas(5) = "signature='" & Trim(selrrequisitnrs("freportby")) & "'"
            End If
        '********** 09/05/11 **********
            
        CrystalReport4.WindowState = crptMaximized
        CrystalReport4.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport4.Action = 1
    
    'End If
    '''
       
    If Trim(SELDHEMCYTO_RS("FLD1")) = "YES" Then
    
    End If
 
 '''  NEW
 If Trim(SELDHEMCYTO_RS("CYTO")) = "YES" Or Trim(SELDHEMCYTO_RS("SMAR")) = "YES" Then
   
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"

    
    
'     If Trim(txt_modcode.Text) = "HL" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShlN1A.RPT"
         
         For i = 0 To 25
            CrystalReport1.Formulas(i) = ""
         Next i
         
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
            str1 = ""
'            STR1 = "Entered By " & Trim(cmb_enterby.Text)
'            If Len(Trim(cmb_enterby.Text)) > 0 Then
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            Else
'                 CRYSTALREPORT1.Formulas(9) = "ENT = '" & "" & "'"
'            End If
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
           
           
           CrystalReport1.WindowState = crptMaximized
           CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
           CrystalReport1.Action = 1
    ' End If
 End If  'CYTO/SURFACE MARKER
 

 
 
 ''   NEW
 
 If REP1 = True And REP2 = True Then
              MsgBox "There are two reports print each one of them"
 End If
End If

 '' HL ADDENDUM
'      If ADDENDUMFALG = True Then
'        CrystalReport2.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        CRYSTALREPORT1.Formulas(1) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
'        CrystalReport2.ReportFileName = "c:\DISREPORTSEC\DISHLADD.RPT"
'        CrystalReport2.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'        CrystalReport2.WindowState = crptMaximized
'           CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'           CrystalReport2.Action = 1
'        MsgBox "Print also the Addendum Report"
'       End If

If opt_staff.Value = True Then
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"

 
'     If Trim(txt_modcode.Text) = "EK" Then
'            CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISEHLASTAFF.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
''        If Trim(prepath) = "" Then
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
''        Else
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
''        End If
'        CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
'        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
'            STR1 = ""
'            STR1 = "Entered By " & Trim(cmb_enterby.Text)
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'         Else
'
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'    End If
'  If Trim(txt_modcode.Text) = "HL" Then
'' NEW CODES FOR STAFF REPORTS
  
  If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
   
   '***** Shraddha 09/03/11 *****
        'str = "SELECT * FROM DISLIB.HINTERFACE WHERE REQNO=?"
        SEL_DHINTER_QR(0) = VAR_REQNNO
        Set SEL_DHINTER_RS = SEL_DHINTER_QR.OpenResultset(1, 3)
        If Not SEL_DHINTER_RS.EOF Then
        '*****************************
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
        
        'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' "
        CrystalReport1.SelectionFormula = "{reporthlf1.REQNO} = '" & VAR_REQNNO & "' "
        'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISAUTOCBC.RPT"
        'Commented and added by Shraddha on 13/09/10
        CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISAUTOCBC.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "F" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
            str1 = ""
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        
        '***** Shraddha 04/01/11 *****
                    
            sel_dhemcbc_qr(0) = Trim(VAR_REQNNO)
            sel_dhemcbc_qr(1) = "R"
            Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 1)
            
            'Dim strRBC, strWBC, strPLATELET As String
            strRBC = ""
            strWBC = ""
            strPLATELET = ""
            While Not sel_dhemcbc_rs.EOF
                'strRBC = strRBC & Trim(SEL_DHEMCBC_RS("PRO_CODE")) & ":" & Trim(SEL_DHEMCBC_RS("OBSERVALUE")) & " "
                strRBC = strRBC & Trim(sel_dhemcbc_rs("TESTNAME")) & ":" & Trim(sel_dhemcbc_rs("OBSERVALUE")) & " "
                sel_dhemcbc_rs.MoveNext
            Wend
            
            CrystalReport1.Formulas(13) = "RBC = '" & strRBC & "'"
            
            '*****
            sel_dhemcbc_qr(0) = Trim(VAR_REQNNO)
            sel_dhemcbc_qr(1) = "W"
            Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 1)
                        
            While Not sel_dhemcbc_rs.EOF
                'strWBC = strWBC & Trim(SEL_DHEMCBC_RS("PRO_CODE")) & ":" & Trim(SEL_DHEMCBC_RS("OBSERVALUE")) & " "
                strWBC = strWBC & Trim(sel_dhemcbc_rs("TESTNAME")) & ":" & Trim(sel_dhemcbc_rs("OBSERVALUE")) & " "
                sel_dhemcbc_rs.MoveNext
            Wend
            
            CrystalReport1.Formulas(14) = "WBC = '" & strWBC & "'"
            
            '*****
            sel_dhemcbc_qr(0) = Trim(VAR_REQNNO)
            sel_dhemcbc_qr(1) = "P"
            Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 1)
                        
            While Not sel_dhemcbc_rs.EOF
                'strPLATELET = strPLATELET & Trim(SEL_DHEMCBC_RS("PRO_CODE")) & ":" & Trim(SEL_DHEMCBC_RS("OBSERVALUE")) & " "
                strPLATELET = strPLATELET & Trim(sel_dhemcbc_rs("TESTNAME")) & ":" & Trim(sel_dhemcbc_rs("OBSERVALUE")) & " "
                sel_dhemcbc_rs.MoveNext
            Wend
            
            CrystalReport1.Formulas(15) = "PLATELET = '" & strPLATELET & "'"
            
            '***** Staff *****
            CrystalReport1.SelectionFormula = "{REPORTHLF1.REQNO} = '" & Trim(VAR_REQNNO) & "'"
            CrystalReport1.ReportFileName = "C:\DisreportsecLF\DISAUTOCBC.RPT"
            
            SELSTAFFINFO_QR(0) = Trim(VAR_REQNNO)
            Set SELSTAFFINFO_RS = SELSTAFFINFO_QR.OpenResultset(2, 3)
            CrystalReport1.Formulas(2) = "NAM = '" & Trim(SELSTAFFINFO_RS("pname")) & "'"
            
            CrystalReport1.Formulas(3) = "staffagesex = '" & Trim(SELSTAFFINFO_RS("AGE")) & "/" & Trim(SELSTAFFINFO_RS("SEX")) & "'"
            
            '*****************
            
        
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
        '***************************** 09/03911
   Else
   
   
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
      
'     If Trim(txt_modcode.Text) = "HL" Then
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"    'AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
'        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShlN1AASTAFF.RPT"
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISAUTOCBCS.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
            str1 = ""
         Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        'CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
        'CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
        'PRINT REPORT
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
     'End If
 'End If  'CBC/HEM/CG
 
 
 
   ' If Trim(SELDHEMCYTO_RS("HEM")) = "YES" Then
        
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
        'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
        
'        If Trim(txt_modcode.Text) = "HL" Then
            CrystalReport4.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' "
        
            CrystalReport4.ReportFileName = "C:\DISREPORTSEC\HEMATOSTAFF.RPT"
        
            CrystalReport4.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            'CrystalReport4.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
            
            CrystalReport4.WindowState = crptMaximized
            CrystalReport4.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
            CrystalReport4.Action = 1
       ' End If
    
   ' End If
 
 
 '''  NEW
 If Trim(SELDHEMCYTO_RS("CYTO")) = "YES" Or Trim(SELDHEMCYTO_RS("SMAR")) = "YES" Then
   
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"

    
    
     'If Trim(txt_modcode.Text) = "HL" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShlN1ASTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
            str1 = ""
'            STR1 = "Entered By " & Trim(cmb_enterby.Text)
'            If Len(Trim(cmb_enterby.Text)) > 0 Then
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            Else
'                 CRYSTALREPORT1.Formulas(9) = "ENT = '" & "" & "'"
'            End If
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
           
          If Trim(txt_lapscore.Text) = "" And Trim(TXT_MPO.Text) = "" And Trim(TXT_SBB.Text) = "" And Trim(TXT_PAS.Text) = "" And Trim(TXT_NSE.Text) = "" And Trim(TXT_TRAP.Text) = "" Then
           'If CYTOCHEMFLAG = True Then
             CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
           Else
            CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
           'End If
           End If
'            If Trim(TXT_BT.Text) = "" And Trim(TXT_PT(0).Text) = "" And Trim(TXT_PLATELETSCP.Text) = "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_TT.Text) = "" And Trim(TXT_FIBRINOGEN.Text) = "" And Trim(TXT_FDPDMIER.Text) = "" And Trim(txt_others.Text) = "" Then
'
'             CRYSTALREPORT1.Formulas(11) = "CP = '" & "" & "'"
'           Else
'            CRYSTALREPORT1.Formulas(11) = "CP = '" & "ABC" & "'"
'
'           End If

           If Trim(TXT_NORMOBLASTS.Text) = "" And Trim(TXT_OTHERCELLSRBC.Text) = "" And Trim(txt_rbcmorphology.Text) = "" And Trim(txt_reticcount.Text) = "" And Trim(txt_malarialparasite.Text) = "" And Trim(TXT_ESRRBC.Text) = "" Then

             CrystalReport1.Formulas(15) = "RBC = '" & "" & "'"
           Else
            CrystalReport1.Formulas(15) = "RBC = '" & "ABC" & "'"

           End If
           If Trim(txt_morphology.Text) = "" And Trim(TXT_ADEQUACY.Text) = "" Then

             CrystalReport1.Formulas(16) = "PLATELETS = '" & "" & "'"
           Else
            CrystalReport1.Formulas(16) = "PLATELETS = '" & "ABC" & "'"

           End If
           If Trim(txt_prolymphocytes.Text) = "" And Trim(txt_Sezarycells.Text) = "" And Trim(txt_hairycells.Text) = "" And Trim(txt_othercells.Text) = "" Then
                CrystalReport1.Formulas(17) = "prolyn = '" & "" & "'"
           Else
                CrystalReport1.Formulas(17) = "prolyn = '" & "123" & "'"
           End If
           If Trim(txt_blasts.Text) = "" And Trim(TXT_PROMYELOCYTESWBC.Text) = "" And Trim(txt_myelocytes.Text) = "" And Trim(txt_metamyelocytes.Text) = "" And Trim(txt_plasmacells.Text) = "" And Trim(txt_alymphocytes.Text) = "" Then
                CrystalReport1.Formulas(18) = "blast = '" & "" & "'"
           Else
                CrystalReport1.Formulas(18) = "blast = '" & "123" & "'"
           End If
           If Trim(txt_mneutrophils.Text) = "" And Trim(txt_mlymphocytes.Text) = "" And Trim(txt_mmonocytes.Text) = "" And Trim(txt_meosinophils.Text) = "" And Trim(txt_mbasophills.Text) = "" Then
                CrystalReport1.Formulas(19) = "cbcm = '" & "" & "'"
           Else
                CrystalReport1.Formulas(19) = "cbcm = '" & "123" & "'"
           End If
           If Trim(TXT_HCT.Text) = "" And Trim(TXT_MCH.Text) = "" And Trim(TXT_MCV.Text) = "" And Trim(TXT_MCHC.Text) = "" And Trim(TXT_MPV.Text) = "" And Trim(TXT_RDW.Text) = "" And Trim(TXT_ABNORMAL.Text) = "" Then
                CrystalReport1.Formulas(20) = "rbci = '" & "" & "'"
           Else
                CrystalReport1.Formulas(20) = "rbci = '" & "123" & "'"
           End If
           If Trim(txt_neutrophils.Text) = "" And Trim(txt_lymphocytes.Text) = "" And Trim(txt_monocytes.Text) = "" And Trim(txt_eosinophils.Text) = "" And Trim(txt_basophils.Text) = "" And Trim(txt_immaturegranulocytes.Text) = "" Then
                CrystalReport1.Formulas(21) = "dlc = '" & "" & "'"
           Else
                CrystalReport1.Formulas(21) = "dlc = '" & "123" & "'"
           End If
           If Trim(TXT_HB.Text) = "" And Trim(TXT_RBC.Text) = "" And Trim(TXT_TWBC.Text) = "" And Trim(txt_platelets.Text) = "" Then
                CrystalReport1.Formulas(22) = "cbcc = '" & "" & "'"
           Else
                CrystalReport1.Formulas(22) = "cbcc = '" & "123" & "'"
           End If
           ' PRINT REPPORT
           
           CrystalReport1.WindowState = crptMaximized
           CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
           CrystalReport1.Action = 1
     End If
 End If   'CYTO/SURFACE MARKER
 End If '09/03/11
    
    
'' NEW CODES FOR STAFF REPORTS
'  OLD CODE FOR STAF REPORT FOR HL BEGINS

    
'        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & Trim(VAR_REQNNO) & "'"
'        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShlSTAFF.RPT"
'        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
''        If Trim(prepath) = "" Then
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
''        Else
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
''        End If
'        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
'        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
'            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
'            STR1 = ""
'            STR1 = "Entered By " & Trim(cmb_enterby.Text)
'            If Len(Trim(cmb_enterby.Text)) > 0 Then
'                CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'            Else
'                 CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
'            End If
'        Else
'            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
'            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
'
'        If Trim(txt_lapscore.Text) = "" And Trim(TXT_MPO.Text) = "" And Trim(TXT_SBB.Text) = "" And Trim(TXT_PAS.Text) = "" And Trim(TXT_NSE.Text) = "" And Trim(TXT_TRAP.Text) = "" Then
'               'If CYTOCHEMFLAG = True Then
'               CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
'        Else
'              CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
'               'End If
'        End If
'        If Trim(TXT_BT.Text) = "" And Trim(TXT_PT(0).Text) = "" And Trim(TXT_PLATELETSCP.Text) = "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_TT.Text) = "" And Trim(TXT_FIBRINOGEN.Text) = "" And Trim(TXT_FDPDMIER.Text) = "" And Trim(txt_others.Text) = "" Then
'            CrystalReport1.Formulas(11) = "CP = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(11) = "CP = '" & "ABC" & "'"
'        End If
'
'        If Trim(TXT_NORMOBLASTS.Text) = "" And Trim(TXT_OTHERCELLSRBC.Text) = "" And Trim(txt_rbcmorphology.Text) = "" And Trim(txt_reticcount.Text) = "" And Trim(txt_malarialparasite.Text) = "" And Trim(TXT_ESRRBC.Text) = "" Then
'             CrystalReport1.Formulas(15) = "RBC = '" & "" & "'"
'        Else
'               CrystalReport1.Formulas(15) = "RBC = '" & "ABC" & "'"
'        End If
'        If Trim(TXT_MORPHOLOGY.Text) = "" And Trim(TXT_ADEQUACY.Text) = "" Then
'            CrystalReport1.Formulas(16) = "PLATELETS = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(16) = "PLATELETS = '" & "ABC" & "'"
'        End If
'        If Trim(txt_prolymphocytes.Text) = "" And Trim(txt_Sezarycells.Text) = "" And Trim(txt_hairycells.Text) = "" And Trim(txt_othercells.Text) = "" Then
'            CrystalReport1.Formulas(17) = "prolyn = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(17) = "prolyn = '" & "123" & "'"
'        End If
'        If Trim(txt_blasts.Text) = "" And Trim(TXT_PROMYELOCYTESWBC.Text) = "" And Trim(txt_myelocytes.Text) = "" And Trim(txt_metamyelocytes.Text) = "" And Trim(txt_plasmacells.Text) = "" And Trim(txt_alymphocytes.Text) = "" Then
'            CrystalReport1.Formulas(18) = "blast = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(18) = "blast = '" & "123" & "'"
'        End If
'        If Trim(txt_mneutrophils.Text) = "" And Trim(txt_mlymphocytes.Text) = "" And Trim(txt_mmonocytes.Text) = "" And Trim(txt_meosinophils.Text) = "" And Trim(txt_mbasophills.Text) = "" Then
'            CrystalReport1.Formulas(19) = "cbcm = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(19) = "cbcm = '" & "123" & "'"
'        End If
'        If Trim(TXT_HCT.Text) = "" And Trim(TXT_MCH.Text) = "" And Trim(TXT_MCV.Text) = "" And Trim(TXT_MCHC.Text) = "" And Trim(TXT_MPV.Text) = "" And Trim(TXT_RDW.Text) = "" And Trim(TXT_ABNORMAL.Text) = "" Then
'            CrystalReport1.Formulas(20) = "rbci = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(20) = "rbci = '" & "123" & "'"
'        End If
'        If Trim(txt_neutrophils.Text) = "" And Trim(txt_lymphocytes.Text) = "" And Trim(txt_monocytes.Text) = "" And Trim(txt_eosinophils.Text) = "" And Trim(txt_basophils.Text) = "" And Trim(txt_immaturegranulocytes.Text) = "" Then
'            CrystalReport1.Formulas(21) = "dlc = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(21) = "dlc = '" & "123" & "'"
'        End If
'        If Trim(txt_hb.Text) = "" And Trim(TXT_RBC.Text) = "" And Trim(TXT_TWBC.Text) = "" And Trim(txt_platelets.Text) = "" Then
'            CrystalReport1.Formulas(22) = "cbcc = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(22) = "cbcc = '" & "123" & "'"
'        End If
'  OLD CODE FOR STAF REPORT FOR HL ENDS
     End If
         
 If REP1 = True And REP2 = True Then
              MsgBox "There are two reports see  each one of them"
 End If
'End If


'***** STAFF
'CrystalReport1.WindowState = crptMaximized
'CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'CrystalReport1.Action = 1
Exit Sub
prtmsg: MsgBox Error(Err)
'txt_msg.Text = ""
'btnenabl ("print")
Exit Sub
MSG:
MsgBox Error(Err)
End Sub
