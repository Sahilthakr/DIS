VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form disw 
   BackColor       =   &H00C0C0C0&
   Caption         =   "DIAGNOSTIC  INFORMATION SYSTEM- WORKLIST"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   13320
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CHK_EDIT 
      BackColor       =   &H80000009&
      Height          =   225
      Left            =   9720
      TabIndex        =   38
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
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
      Left            =   10560
      TabIndex        =   36
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   480
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
      Left            =   10560
      TabIndex        =   33
      Top             =   480
      Width           =   1095
   End
   Begin VB.CheckBox chk_commit 
      BackColor       =   &H80000009&
      Height          =   225
      Left            =   8250
      TabIndex        =   32
      Top             =   4395
      Visible         =   0   'False
      Width           =   855
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   0
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1695
      Left            =   2055
      TabIndex        =   27
      Top             =   4695
      Visible         =   0   'False
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   2990
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "SR.NO."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "CASE NO."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "REQUISITION NO."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "PATH NO."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "SAMPLE REC. DATE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "REPORT DATE"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   11655
      Begin VB.ComboBox Cmb_ReqNo 
         Height          =   315
         Left            =   7065
         TabIndex        =   31
         Top             =   1890
         Width           =   3135
      End
      Begin MSMask.MaskEdBox Msk_CCNo 
         Height          =   375
         Left            =   3240
         TabIndex        =   29
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         HideSelection   =   0   'False
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
         Mask            =   ">??/#####"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton Opt_Caseno 
         Caption         =   "Case No"
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
         Left            =   1560
         TabIndex        =   28
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1335
         Begin VB.OptionButton Opt_Staff 
            Caption         =   "Staff"
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
            TabIndex        =   26
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Opt_Pat 
            Caption         =   "Patient"
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
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.ComboBox cbo_modality 
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
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   675
      End
      Begin VB.OptionButton Opt_case 
         Caption         =   "From Date "
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
         Left            =   1560
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Opt_all 
         Caption         =   "All"
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
         Left            =   1560
         TabIndex        =   6
         Top             =   1560
         Width           =   645
      End
      Begin VB.ComboBox CMB_DEPT 
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
         Left            =   3240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox CMB_DOCTOR 
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
         Left            =   7080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin MSMask.MaskEdBox MaskEd_to_date 
         Height          =   375
         Left            =   7080
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         AutoTab         =   -1  'True
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
         Left            =   3240
         TabIndex        =   4
         Top             =   1200
         Width           =   1245
         _ExtentX        =   2196
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
      Begin VB.TextBox txt_doctor 
         Height          =   285
         Left            =   7080
         TabIndex        =   20
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txt_dept 
         Height          =   285
         Left            =   3600
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Requisitn No."
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
         Left            =   5280
         TabIndex        =   30
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2040
         TabIndex        =   19
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "To Date "
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
         Left            =   5640
         TabIndex        =   18
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label lbl_Dept 
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl_doctor 
         Caption         =   "Doctor"
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
         Left            =   5640
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   240
      TabIndex        =   11
      Top             =   7680
      Width           =   11655
      Begin VB.CommandButton cmd_edit 
         Caption         =   "&EDIT"
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
         Left            =   4680
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CMD_COMMIT 
         Caption         =   "&COMMIT"
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
         Left            =   5520
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton CMD_VIEW 
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
         Left            =   2520
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton CMD_CANCEL 
         Caption         =   "&CANCEL"
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
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton CMD_PRINT 
         Caption         =   "&PRINT"
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
         Left            =   8520
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton CMD_BACK 
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
         Height          =   375
         Left            =   7080
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf_reqstat 
      Height          =   2775
      Left            =   555
      TabIndex        =   8
      ToolTipText     =   "CLICK ON THE REQUISTION NO OR CASENO."
      Top             =   4050
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   10
      ForeColor       =   0
      ForeColorFixed  =   192
      WordWrap        =   -1  'True
      Enabled         =   -1  'True
      FormatString    =   "SR.NO. |^   PATH NO.   |^           CASE NO. |^   REQUISITION NO.       |^ SAMPLE REC. DATE |^ COMMIT |^ REPORT DATE|^EDIT"
   End
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   240
      TabIndex        =   22
      Top             =   3960
      Width           =   11655
   End
   Begin MSWinsockLib.Winsock WIN_CLIENT 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1080
      TabIndex        =   39
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
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
      TabIndex        =   35
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
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
      TabIndex        =   34
      Top             =   480
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "WAITING FOR FINAL REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3360
      TabIndex        =   9
      Top             =   3840
      Width           =   3195
   End
   Begin VB.Label LBL_DISW 
      BackStyle       =   0  'Transparent
      Caption         =   "DIAGNOSTIC  INFORMATION SYSTEM - WORKLIST"
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
      Left            =   2040
      TabIndex        =   7
      Top             =   720
      Width           =   7335
   End
End
Attribute VB_Name = "disw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SELDBMAREPORT_QR As rdoQuery
Dim SELDBMAREPORT_RS As rdoResultset
Dim SELDBMAREPORT_FLAG As Boolean
Dim SELECREPORTQR As rdoQuery
Dim SELECREPORTRS As rdoResultset
Dim SELECREPORTFLAG As Boolean
Dim REFDR_QR As rdoQuery
Dim REFDR_RS As rdoResultset
Dim SELDHEMCYTO_QR As rdoQuery
Dim SELDHEMCYTO_RS As rdoResultset
Dim SELDBMASUB_QR As rdoQuery
Dim SELDADDENDUMQR As rdoQuery
Dim SELDADDENDUMRS As rdoResultset
Dim selrrequisitnqr As rdoQuery
Dim selrrequisitnrs As rdoResultset
Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim SELRMODALITY_QR As rdoQuery
Dim SELRMODALITY_RS As rdoResultset

Dim rq_grid_det As rdoQuery
Dim rs_grid_det As rdoResultset

'Public WORKLIST As Boolean

Dim SELSTREPORTQR As rdoQuery
Dim SELSTREPORTRS As rdoResultset

Dim UPDCOMMIT_QR As rdoQuery

Dim DMULPROCE_QR As rdoQuery
Dim DMULPROCE_RS As rdoResultset
Dim DMULPROCEFLAG  As Boolean
''GRID DATA
Dim SELPATH_QR As rdoQuery
Dim SELPATH_RS As rdoResultset
Dim SELPATH_FLAG As Boolean

Dim SELDTMNO_QR As rdoQuery
Dim SELDTMNO_RS As rdoResultset
Dim SELDTMNO_FLAG As Boolean

''FOR DISLNKUSR
Dim LNK_QR As rdoQuery
Dim LNK_RS As rdoResultset
Dim LNK_FLAG As Boolean

''DTM BLANK REPORT CHECK
Dim SELDTMTAB_QR As rdoQuery
Dim SELDTMTAB_RS As rdoResultset
Dim SELDTMTAB_FLAG As Boolean

Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset
Dim SELTA101_FLAG As Boolean

Dim SELDREUISITION_QR As rdoQuery
Dim SELDREUISITION_RS As rdoResultset

Dim SELRREUISITION_QR As rdoQuery
Dim SELRREUISITION_RS As rdoResultset
Dim var_caseno As String

Dim SELEMPBASINFO_QR As rdoQuery
Dim SELEMPBASINFO_RS As rdoResultset
Dim SELEMPBASINFOFLAG  As Boolean

Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset

Dim SELMA301_QR As rdoQuery
Dim SELMA301_RS As rdoResultset

Dim SELDFREPORT_QR As rdoQuery
Dim SELDFREPORT_RS As rdoResultset

Dim UPDRREQUISITION_QR As rdoQuery
''GRID DATA

'' FOR CASE NO AND REQUISITN NO

Dim sel_ReqNo_qr As rdoQuery
Dim sel_ReqNo_Rs As rdoResultset
Dim sel_ReqNo_flg As Boolean

Dim sel_GReqNo_qr As rdoQuery
Dim sel_GReqNo_rs As rdoResultset
Dim sel_GReqNo_flg As Boolean
Dim CASREQ_FLG As Boolean

'' FOR CASE NO AND REQUISITN NO

Dim rq_modcd As rdoQuery
Dim rs_modcd As rdoResultset
Dim SELMA106_QR As rdoQuery
Dim SELMA106_RS As rdoResultset
'ALL DEPT
Dim SELMA106_QR_ALL As rdoQuery
Dim SELMA106_RS_ALL As rdoResultset
'ALL DEPT
Dim SELREQN_QR As rdoQuery
Dim SELREQN_RS As rdoResultset
Dim selreqnn_qr As rdoQuery
Dim selreqnn_rs As rdoResultset
Dim SELREQUI_QR As rdoQuery
Dim SELREQUI_RS As rdoResultset
Dim FROMFLAG As Boolean
Dim RISAU21_QR As rdoQuery
Dim RISAU21_RS As rdoResultset
Dim SELRISLNKUSR_QR As rdoQuery
Dim SELRISLNKUSR_RS As rdoResultset



''FOR ALL MODALITY
Dim SELALL_QR As rdoQuery
Dim SELALL_RS As rdoResultset

Dim selreqnn_qr1 As rdoQuery
Dim selreqnn_rs1 As rdoResultset

''FOR ALL MODALITY

''FOR REPORT
Dim str As String
Dim VAR_REQNNO As String

Private Sub chk_commit_Click()
    
    If chk_Commit.Value = 1 Then
      msf_reqstat.TextMatrix(msf_reqstat.row, 5) = "Y"
      chk_Commit.Visible = False
    Else
      chk_Commit.Visible = False
      msf_reqstat.TextMatrix(msf_reqstat.row, 5) = ""
    End If

End Sub
Private Sub CHK_EDIT_Click()
    
If chk_edit.Value = 1 Then
Dim SS As String
 SS = Mid(msf_reqstat.TextMatrix(msf_reqstat.row, 3), 5, 2)
    
    
    WORKLIST = True
 If SS = "BM" Then

        bmareport.show
        bmareport.cmd_view_Click
       ' bmareport.cmb_dept.AddItem "PATHOLOGY$FZZ"
        'bmareport.cmb_dept.Text = "PATHOLOGY$FZZ"
        'bmareport.cmb_dept_Click
        'bmareport.cmb_modality.AddItem "FNAC$FA"
        'bmareport.cmb_modality.Text = "FNAC$FA"
        'bmareport.cmb_modality_Click
        bmareport.MaskEdBox1.Text = msf_reqstat.TextMatrix(msf_reqstat.row, 2)
        bmareport.MaskEdBox1_LostFocus
        bmareport.cmb_reqno.AddItem msf_reqstat.TextMatrix(msf_reqstat.row, 3)
        bmareport.cmb_reqno.Text = msf_reqstat.TextMatrix(msf_reqstat.row, 3)
        bmareport.cmb_reqno_Click
        
        
   ElseIf SS = "HL" Then

        hemcbcreport.show
        hemcbcreport.cmd_view_Click
        hemcbcreport.cmb_dept.AddItem "HAEMATOPATHOLOGY$UZZ"
        hemcbcreport.cmb_dept.Text = "HAEMATOPATHOLOGY$UZZ"
        hemcbcreport.cmb_dept_Click
        hemcbcreport.cmb_modality.AddItem "HEAMATOLOGY$HL"
        hemcbcreport.cmb_modality.Text = "HEAMATOLOGY$HL"
        hemcbcreport.cmb_modality_Click
        hemcbcreport.MaskEdBox1.Text = msf_reqstat.TextMatrix(msf_reqstat.row, 2)
        hemcbcreport.MaskEdBox1_LostFocus
        hemcbcreport.cmb_reqno.AddItem msf_reqstat.TextMatrix(msf_reqstat.row, 3)
        hemcbcreport.cmb_reqno.Text = msf_reqstat.TextMatrix(msf_reqstat.row, 3)
        hemcbcreport.cmb_reqno_Click
        
   ElseIf SS = "BF" Then

        HEMBFREPORTOLD.show
        HEMBFREPORTOLD.cmd_view_Click
        HEMBFREPORTOLD.cmb_dept.AddItem "HAEMATOPATHOLOGY$UZZ"
        HEMBFREPORTOLD.cmb_dept.Text = "HAEMATOPATHOLOGY$UZZ"
        HEMBFREPORTOLD.cmb_dept_Click
        HEMBFREPORTOLD.cmb_modality.AddItem "BODY FLUID REPORT$BF"
        HEMBFREPORTOLD.cmb_modality.Text = "BODY FLUID REPORT$BF"
        HEMBFREPORTOLD.cmb_modality_Click
        HEMBFREPORTOLD.MaskEdBox1.Text = msf_reqstat.TextMatrix(msf_reqstat.row, 2)
        HEMBFREPORTOLD.MaskEdBox1_LostFocus
        HEMBFREPORTOLD.cmb_reqno.AddItem msf_reqstat.TextMatrix(msf_reqstat.row, 3)
        HEMBFREPORTOLD.cmb_reqno.Text = msf_reqstat.TextMatrix(msf_reqstat.row, 3)
        HEMBFREPORTOLD.cmb_reqno_Click
    
   ElseIf SS = "SP" Then

        SPPATHREPORT.show
        SPPATHREPORT.cmd_view_Click
        SPPATHREPORT.cmb_dept.AddItem "PATH0LOGY$FZZ"
        SPPATHREPORT.cmb_dept.Text = "PATH0LOGY$FZZ"
        SPPATHREPORT.cmb_dept_Click
        SPPATHREPORT.cmb_modality.AddItem "HISTOPATHOLOGY$SP"
        SPPATHREPORT.cmb_modality.Text = "HISTOPATHOLOGY$SP"
        SPPATHREPORT.cmb_modality_Click
        SPPATHREPORT.MaskEdBox1.Text = msf_reqstat.TextMatrix(msf_reqstat.row, 2)
        SPPATHREPORT.MaskEdBox1_LostFocus
        SPPATHREPORT.cmb_reqno.AddItem msf_reqstat.TextMatrix(msf_reqstat.row, 3)
        SPPATHREPORT.cmb_reqno.Text = msf_reqstat.TextMatrix(msf_reqstat.row, 3)
        SPPATHREPORT.cmb_reqno_Click
 msf_reqstat.TextMatrix(msf_reqstat.row, 7) = "Y"
    chk_edit.Visible = False
    'QRCLOSE
    Me.Hide
    ElseIf SS = "FA" Then

        finnedaspcyt_report.show
        finnedaspcyt_report.cmd_view_Click
        finnedaspcyt_report.cmb_dept.AddItem "PATHOLOGY$FZZ"
        finnedaspcyt_report.cmb_dept.Text = "PATHOLOGY$FZZ"
        finnedaspcyt_report.cmb_dept_Click
        finnedaspcyt_report.cmb_modality.AddItem "FNAC$FA"
        finnedaspcyt_report.cmb_modality.Text = "FNAC$FA"
        finnedaspcyt_report.cmb_modality_Click
        finnedaspcyt_report.MaskEdBox1.Text = msf_reqstat.TextMatrix(msf_reqstat.row, 2)
        finnedaspcyt_report.MaskEdBox1_LostFocus
        finnedaspcyt_report.cmb_reqno.AddItem msf_reqstat.TextMatrix(msf_reqstat.row, 3)
        finnedaspcyt_report.cmb_reqno.Text = msf_reqstat.TextMatrix(msf_reqstat.row, 3)
        finnedaspcyt_report.cmb_reqno_Click
'
'     ElseIf SS = "CN" Then
'
'        cytology_report.show
'      '  cytology_report.cmd_view_Click
'        cytology_report.cmb_dept.AddItem "PATHOLOGY$FZZ"
'        cytology_report.cmb_dept.Text = "PATHOLOGY$FZZ"
'      ' cytology_report.cmb_dept_Click
'        cytology_report.cmb_modality.AddItem "CYTOLOGY NON GYNAEC$CN"
'        cytology_report.cmb_modality.Text = "CYTOLOGY NON GYNAEC$CN"
'      '  cytology_report.cmb_modality_Click
'        cytology_report.MaskEdBox1.Text = msf_reqstat.TextMatrix(msf_reqstat.row, 2)
'      '  cytology_report.MaskEdBox1_LostFocus
'        cytology_report.cmb_reqno.AddItem msf_reqstat.TextMatrix(msf_reqstat.row, 3)
'        cytology_report.cmb_reqno.Text = msf_reqstat.TextMatrix(msf_reqstat.row, 3)
'    '  cytology_report.cmb_reqno_Click


 msf_reqstat.TextMatrix(msf_reqstat.row, 7) = "Y"
    chk_edit.Visible = False
    'QRCLOSE
    Me.Hide
    End If

   
Else
    chk_edit.Visible = False
    msf_reqstat.TextMatrix(msf_reqstat.row, 7) = ""
End If
    
'    Unload Me
End Sub

''FOR REPORT
Private Sub CMB_DOCTOR_LostFocus()
'cbo_modality.SetFocus
End Sub

Private Sub cmb_reqno_Click()
    Dim J1 As Integer
    Dim exst_flg As Boolean
    
    exst_flg = False
    If Trim(cmb_reqno.Text) = "" Then
        MsgBox ("Select Any of the Requisitn No."), vbInformation
        Exit Sub
    End If
    i = 1
    msf_reqstat.Rows = 3
    
    If Trim(cmb_reqno.Text) = "ALL" Then
        i = 1
        msf_reqstat.Rows = 1
        sel_ReqNo_qr(0) = Trim(Msk_CCNo.FormattedText)
        If Trim(USRGRP) = "SYSADM" Or Trim(USRGRP) = "HOD" Or Trim(USRGRP) = "MEDADMN" Then
            sel_ReqNo_qr(1) = Trim(cmb_doctor.Text)
        Else
            sel_ReqNo_qr(1) = Trim(txt_doctor.Text)
        End If
        If sel_ReqNo_flg = False Then
            Set sel_ReqNo_Rs = sel_ReqNo_qr.OpenResultset(1, 3)
            sel_ReqNo_flg = True
        Else
            sel_ReqNo_Rs.MoreResults
            sel_ReqNo_Rs.Requery
        End If
       
        If sel_ReqNo_Rs.RowCount > 0 Then
            sel_ReqNo_Rs.MoveFirst
            While Not sel_ReqNo_Rs.EOF
                SELPATH_QR(0) = Trim(sel_ReqNo_Rs("reqnno"))
                Set SELPATH_RS = SELPATH_QR.OpenResultset(1, 3)
                msf_reqstat.AddItem i & vbTab & Trim(SELPATH_RS("PATHNO")) & vbTab & Trim(sel_ReqNo_Rs("caseno")) & vbTab & req_chglen(sel_ReqNo_Rs("reqnno")) & vbTab & Trim(sel_ReqNo_Rs("APNTDT")) & vbTab & vbTab & Trim(sel_ReqNo_Rs("freportdt"))
                'msf_reqstat.AddItem i & vbTab & Trim(sel_ReqNo_Rs("caseno")) & vbTab & req_chglen(sel_ReqNo_Rs("reqnno")) & vbTab & vbTab & Trim(sel_GReqNo_rs("APNTDT")) & vbTab & Trim(sel_GReqNo_rs("freportdt"))
                sel_ReqNo_Rs.MoveNext
                i = i + 1
            Wend
            CASREQ_FLG = True
        End If
    Else
        
        If Trim(req_chglen(msf_reqstat.TextMatrix(1, 3))) = "" Then
               i = 1
                msf_reqstat.Rows = 1
                sel_GReqNo_qr(0) = Trim(cmb_reqno.Text)
                
                If sel_GReqNo_flg = False Then
                    Set sel_GReqNo_rs = sel_GReqNo_qr.OpenResultset(1, 3)
                    sel_GReqNo_flg = True
                Else
                    sel_GReqNo_rs.MoreResults
                    sel_GReqNo_rs.Requery
                End If
                
                SELPATH_QR(0) = Trim(cmb_reqno.Text)
                Set SELPATH_RS = SELPATH_QR.OpenResultset(1, 3)
                msf_reqstat.AddItem i & vbTab & Trim(SELPATH_RS("PATHNO")) & vbTab & Trim(sel_GReqNo_rs("caseno")) & vbTab & req_chglen(sel_GReqNo_rs("reqnno")) & vbTab & Trim(sel_GReqNo_rs("APNTDT")) & vbTab & vbTab & Trim(sel_GReqNo_rs("freportdt"))
                Exit Sub
            End If
            
            ' for checking existing values
        For J1 = 1 To msf_reqstat.Rows - 1
                        
            If Trim(req_chglen(msf_reqstat.TextMatrix(J1, 3))) = Trim(cmb_reqno.Text) Then
                MsgBox ("Requisitn No Already Added to the Grid"), vbInformation
                exst_flg = True
                Exit Sub
            End If
             
        Next J1
        
        ' for inserting new values
        If exst_flg = False Then
                i = msf_reqstat.row + 1
                'msf_reqstat.Rows = msf_reqstat.Rows + 1
                 
                sel_GReqNo_qr(0) = Trim(cmb_reqno.Text)
                If sel_GReqNo_flg = False Then
                    Set sel_GReqNo_rs = sel_GReqNo_qr.OpenResultset(1, 3)
                    sel_GReqNo_flg = True
                Else
                    sel_GReqNo_rs.MoreResults
                    sel_GReqNo_rs.Requery
                End If
                
                SELPATH_QR(0) = Trim(cmb_reqno.Text)
                Set SELPATH_RS = SELPATH_QR.OpenResultset(1, 3)
                msf_reqstat.AddItem i & vbTab & Trim(SELPATH_RS("PATHNO")) & vbTab & Trim(sel_GReqNo_rs("caseno")) & vbTab & req_chglen(sel_GReqNo_rs("reqnno")) & vbTab & Trim(sel_GReqNo_rs("APNTDT")) & vbTab & vbTab & Trim(sel_GReqNo_rs("freportdt"))
                Exit Sub
 
        End If
    End If
    
    
End Sub

Private Sub cmd_back_Click()
Unload Me
DIS_MAIN.show
WORKLIST = False
End Sub

Private Sub cmd_cancel_Click()
cmd_view.Enabled = True
cmd_print.Enabled = False
cmdview
'CMDPRINT
cmd_back.Enabled = True
cmd_cancel.Enabled = False

'CMB_DEPT.ABSOLUTPOSITION = -1
'CMB_DOCTOR.Text = ""
cbo_modality.Text = ""
OPT_ALL.Enabled = False
Opt_case.Enabled = False
MaskEd_from_date.Mask = "##/##/####"
MaskEd_from_date.Text = "__/__/____"
MaskEd_to_date.Mask = "##/##/####"
MaskEd_to_date.Text = "__/__/____"

Msk_CCNo.Mask = ">??/#####"
Msk_CCNo.Text = "__/_____"
cmb_reqno.Text = ""
OPT_CASENO.Enabled = False
Msk_CCNo.Visible = False
Label1.Visible = True
cmb_reqno.Visible = False
Opt_Pat.Value = False
opt_staff.Value = False
chk_Commit.Visible = False
chk_edit.Visible = False
cmb_dept.Enabled = False
cmb_doctor.Enabled = False
cbo_modality.Enabled = False
OPT_ALL.Enabled = False
Opt_case.Enabled = False
MaskEd_from_date.Enabled = False
MaskEd_to_date.Enabled = False

msf_reqstat.Rows = 1

End Sub

Private Sub cmd_commit_Click()
Dim count, i As Integer
count = 0

For i = 1 To msf_reqstat.Rows - 1
    VAR_REQNNO = req_chglen(msf_reqstat.TextMatrix(i, 3))
    If Trim(VAR_REQNNO) <> "" And msf_reqstat.TextMatrix(i, 5) = "Y" Then
        UPDRREQUISITION_QR(0) = VAR_REQNNO
        UPDRREQUISITION_QR.Execute
        
        UPDCOMMIT_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY")) 'Format(txt_date.Text, "YYYY-MM-DD")
        UPDCOMMIT_QR(1) = Format(txt_time.Text, "HH:MM:SS")
        UPDCOMMIT_QR(2) = req_chglen(msf_reqstat.TextMatrix(i, 3)) 'req_chglen(cmb_reqno.Text)
        UPDCOMMIT_QR.Execute
        count = count + 1
        sendRMSG req_chglen(msf_reqstat.TextMatrix(i, 3)), msf_reqstat.TextMatrix(i, 2), "F", WIN_CLIENT.LocalIP
    End If
Next i
MsgBox count & " Report is Commited Successfully !!", vbInformation, "Work List"
    
    
'''VAR_REQNNO = req_chglen(msf_reqstat.TextMatrix(msf_reqstat.Row, 3))  '''req_chglen(mask_edit_req.Text)
''''var_caseno = msf_reqstat.TextMatrix(msf_reqstat.Row, 1)
'''If MsgBox("U'r Going to Commit a Report of Path No. " & Trim(msf_reqstat.TextMatrix(msf_reqstat.Row, 2)) & ", is it OK for U ?", vbOKCancel) = vbOK Then
'''    If Trim(VAR_REQNNO) <> "" Then
'''        UPDRREQUISITION_QR(0) = VAR_REQNNO
'''        UPDRREQUISITION_QR.Execute
'''        MsgBox " Report is Commited Successfully !!", vbInformation, "Work List"
'''    End If
    If CASREQ_FLG = False Then
        'msf_reqstat.clear
        Opt_all_LostFocus
        MaskEd_to_date_LostFocus
    Else
        msf_reqstat.clear
        cmb_reqno.Text = "ALL"
        cmb_reqno_Click
    End If


Exit Sub
End Sub

Private Sub cmd_edit_Click()

    
Me.Hide
WORKLIST = True
SPPATHREPORT.show


'Dim count, i As Integer
'count = 0
'
'For i = 1 To msf_reqstat.Rows - 1
'    VAR_REQNNO = req_chglen(msf_reqstat.TextMatrix(i, 3))
'    If Trim(VAR_REQNNO) <> "" And msf_reqstat.TextMatrix(i, 5) = "Y" Then
'        UPDRREQUISITION_QR(0) = VAR_REQNNO
'        UPDRREQUISITION_QR.Execute
'
'        UPDCOMMIT_QR(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY")) 'Format(txt_date.Text, "YYYY-MM-DD")
'        UPDCOMMIT_QR(1) = Format(txt_time.Text, "HH:MM:SS")
'        UPDCOMMIT_QR(2) = req_chglen(Cmb_ReqNo.Text)
'        UPDCOMMIT_QR.Execute
'        count = count + 1
'    End If
'Next i
'MsgBox count & " Report is Commited Successfully !!", vbInformation, "Work List"
'
    

    If CASREQ_FLG = False Then
        'msf_reqstat.clear
        Opt_all_LostFocus
        MaskEd_to_date_LostFocus
    Else
        msf_reqstat.clear
        cmb_reqno.Text = "ALL"
        cmb_reqno_Click
    End If


Exit Sub

End Sub

Private Sub cmd_print_Click()
'CMD_VIEW.Enabled = True
cmd_print.Enabled = False
'CMDPRINT
cmdview
cmd_back.Enabled = True
cmd_cancel.Enabled = True
On Error GoTo MSG:
Dim var_yr_f As Integer, var_yr_t As Integer
Dim var_mon_f As Integer, var_mon_t As Integer
Dim var_day_f As Integer, var_day_t As Integer
Dim var_date_f As String, var_date_t As String
   
If FROMFLAG = True Then
    FROMFLAG = False
    var_date_f = Year(MaskEd_from_date) & "," & month(MaskEd_from_date) & "," & Day(MaskEd_from_date)
    var_date_t = Year(MaskEd_to_date) & "," & month(MaskEd_to_date) & "," & Day(MaskEd_to_date)

    CR1.ReportFileName = "C:\DISREPORTSec\DISW2.RPT"
    'Cr_reqstat.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & Trim(cbo_modality) & " ' AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
    CR1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ")   AND {RREQUISITN.MODALITYCD}='" & Trim(cbo_modality.Text) & "'" & "and" & "{rrequisitn.freportby}= '" & txt_doctor.Text & "'" & "and" & "{rrequisitn.reqnstat}='T'"
    CR1.Formulas(0) = "DEPT = '" & cmb_dept.Text & "'"
    CR1.Formulas(1) = "DOCTOR = '" & cmb_doctor.Text & "'"
    CR1.Formulas(2) = "MOD = '" & cbo_modality.Text & "'"
    CR1.Formulas(3) = "FROMDT = '" & MaskEd_from_date.Text & "'"
    CR1.Formulas(4) = "TODT = '" & MaskEd_to_date.Text & "'"
    CR1.WindowState = crptMaximized
    CR1.Action = 1
Exit Sub

Else
    CR1.ReportFileName = "C:\RISREPORTS\DISW1.RPT"
    'Cr_reqstat.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & Trim(cbo_modality) & " ' AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
    CR1.SelectionFormula = "{RREQUISITN.MODALITYCD}='" & Trim(cbo_modality.Text) & "'" & "and" & "{rrequisitn.freportby}= '" & cmb_doctor.Text & "'" & "and" & "{rrequisitn.reqnstat}='T'"
    CR1.Formulas(0) = "DEPT = '" & cmb_dept.Text & "'"
    CR1.Formulas(1) = "DOCTOR = '" & cmb_doctor.Text & "'"
    CR1.Formulas(2) = "MOD = '" & cbo_modality.Text & "'"
    CR1.WindowState = crptMaximized
    CR1.Action = 1
Exit Sub
End If
MSG:
MsgBox Error(Err)
Exit Sub
End Sub

Private Sub cmd_view_Click()
cmd_view.Enabled = False
cmd_print.Enabled = True
cmd_back.Enabled = True
cmd_cancel.Enabled = True

cmb_dept.Enabled = True
cmb_doctor.Enabled = True
cbo_modality.Enabled = True
OPT_ALL.Enabled = True
Opt_case.Enabled = True
OPT_ALL.Value = True
Opt_case.Value = False
MaskEd_from_date.Enabled = False
MaskEd_to_date.Enabled = False
cbo_modality.Text = "ALL"
cmb_dept.SetFocus
End Sub

Private Sub Form_Load()
datetime

txt_date = DATE11
txt_time = time1

cmd_view.Enabled = True
cmd_print.Enabled = False
cmd_back.Enabled = True
cmd_cancel.Enabled = False

cmb_dept.Enabled = False
cmb_doctor.Enabled = False
cbo_modality.Enabled = False
OPT_ALL.Enabled = False
Opt_case.Enabled = False
MaskEd_from_date.Enabled = False
MaskEd_to_date.Enabled = False

OPT_CASENO.Enabled = False
Msk_CCNo.Visible = False
Label1.Visible = True
cmb_reqno.Visible = False

datetime
'' FOR CASE NO AND REQUISITN NO
    
   ''DTM BLANK REPORT CHEQUE
    str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=? AND REQNNO=? ORDER BY REQNDT DESC"
    Set REFDR_QR = rdb.CreateQuery("REFDR_QR1HE12", str)
   
    str = "SELECT * FROM DISLIB.DECREPORT WHERE REQNO=?"
    Set SELECREPORTQR = rdb.CreateQuery("SELDECREPORTQRXX12", str)
   
    str = "SELECT * FROM DISLIB.DTMTAB WHERE REQNO=?"
    Set SELDTMTAB_QR = rdb.CreateQuery("SELDTMTAB_STR112", str)
    
    str = "select * from dislib.dhemcyto1 where reqno = ?"
    Set SELDHEMCYTO_QR = rdb.CreateQuery("seledhemcyto11_qr2", str)
    
    str = "select * from Rislib.Rrequisitn where reqnNo=?"
    Set selrrequisitnqr = rdb.CreateQuery("selrrequisitnqr12", str)
    
    str = "select * from dislib.drequisitn where reqno=?"
    Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn1_qr12", str)
    
    str = "SELECT DEPT_CODE FROM DISLIB.DISLNKUSR WHERE USR_ID=?"
    Set LNK_QR = rdb.CreateQuery("L2", str)
    
    str = "SELECT * FROM PATOOLSLIB.TA101 WHERE CAsE_NO=?"
    Set SELTA101_QR = rdb.CreateQuery("RSELTA101_QR12", str)
    
    str = "SELECT * FROM DISLIB.DBMAREPORT WHERE REQNO=?"
    Set SELDBMAREPORT_QR = rdb.CreateQuery("SELDBMAREPORT_STR2", str)
    
    str = "SELECT * FROM DISLIB.DBMASUBREP WHERE REQNO = ?"
    Set SELDBMASUB_QR = rdb.CreateQuery("DBMASUB112", str)
    str = "SELECT pro_code FROM DISLIB.DMULPROCE WHERE REQNO=? "
    Set DMULPROCE_QR = rdb.CreateQuery("DMULPROCE_QR1TUMREPRQ2", str)
    
    str = "SELECT * FROM DISLIB.DADDENDUM WHERE REQNO=?"
    Set SELDADDENDUMQR = rdb.CreateQuery("SELDADDENDUMQR12", str)
    
    str = "SELECT * FROM DISLIB.DSTREPORT WHERE REQNO=?"
    Set SELSTREPORTQR = rdb.CreateQuery("SELSTREPORTQRSVZ2", str)
    
    str = "SELECT * FROM RISLIB.RMODALITY WHERE MODALITYCD=?"
    Set SELRMODALITY_QR = rdb.CreateQuery("SELRMODALITY1_QR12", str)
    SELRMODALITY_QR(0) = "CT"
    Set SELRMODALITY_RS = SELRMODALITY_QR.OpenResultset(2, 3)
    
    str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?, COMREPTIME=? WHERE REQNNO=?"
    Set UPDCOMMIT_QR = rdb.CreateQuery("UPDCOMMIT_QR", str)
    
    ''DTMNO
    str = "SELECT * FROM DISLIB.DTMNO WHERE CASENO=?"
    Set SELDTMNO_QR = rdb.CreateQuery("SELDTMNO_STR122", str)
    
    Set rq_grid_det = rdb.CreateQuery("grid_det2", "select REQNNO,REQNSTAT,APNTDT,MEMONO,RECPTNO,MEMODT,EXAMDT,FREPORTDT,PREPORTDT ,FREPORTBY,RECPTDT,PATNSTAT,recptno,reqndt,ENTERBY from Rislib.RREQUISITN where REQNNO=?")
     rq_grid_det(0) = "aa"
     Set rs_grid_det = rq_grid_det.OpenResultset(2, 3)
    
    str = "select caseno,reqnno,APNTDT,EXAMDT,FREPORTDT from Rislib.Rrequisitn where CASENO=? AND  freportby=? AND REQNSTAT='T'  ORDER BY reqnno"
    Set sel_ReqNo_qr = rdb.CreateQuery("sel_ReqNo_qr", str)
    sel_ReqNo_flg = False
    
    str = "select caseno,reqnno,apntdt,FREPORTDT,EXAMDT from Rislib.Rrequisitn where Reqnno=?"
    Set sel_GReqNo_qr = rdb.CreateQuery("sel_GReqNo_qr", str)
    sel_GReqNo_flg = False
  '' FOR CASE NO AND REQUISITN NO

  FROMFLAG = False
  ''****************GRID DATA********************
  str = "SELECT PATHNO FROM DISLIB.DREQUISITN WHERE REQNO=? ORDER BY PATHNO"
  Set SELPATH_QR = rdb.CreateQuery("SELPATH_STR", str)
  
  str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
  Set SELRREUISITION_QR = rdb.CreateQuery("SELREQ_STR", str)
  
  str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO=?"
  Set SELDREUISITION_QR = rdb.CreateQuery("SELREQ_STR111", str)
  
  str = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
  Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO_QR111009", str)
  
  str = "select * from patoolslib.ta1012 where case_no=?"
  Set selta1012_qr = rdb.CreateQuery("2341selta10121_qr1", str)
    
  str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
  Set SELMA301_QR = rdb.CreateQuery("S1ELMA3011_QR1", str)
  
  str = "SELECT * FROM DISLIB.DFREPORT WHERE REQNO=?"
  Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT1_QR11021", str)
  
  str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT='F',COMMITFLAG='Y' WHERE REQNNO=?"
  Set UPDRREQUISITION_QR = rdb.CreateQuery("UPDRREQUISITION_STR", str)
  
  ''****************GRID DATA********************
''  '' FOR CASE NO AND REQUISITN NO
''
''    STR = "select caseno,reqnno,APNTDT,EXAMDT,FREPORTDT from Rislib.Rrequisitn where CASENO=? AND  freportby=? AND REQNSTAT='T'AND DEPT_CODE='FZZ'"
''    Set sel_ReqNo_qr = rdb.CreateQuery("sel_ReqNo_qr", STR)
''
''    STR = "select caseno,reqnno,apntdt,FREPORTDT,EXAMDT from Rislib.Rrequisitn where Reqnno=?"
''    Set sel_GReqNo_qr = rdb.CreateQuery("sel_GReqNo_qr", STR)
''
''  '' FOR CASE NO AND REQUISITN NO
     If Trim(UID) = "A101370" Then
     str = "select distinct(modalitycd) from rislib.rmodality WHERE DEPT_CODE=? OR DEPT_CODE='UZZ'"
     Else
     str = "select distinct(modalitycd) from rislib.rmodality WHERE DEPT_CODE=?"
     End If
    
  
  Set rq_modcd = rdb.CreateQuery("rq_modcd1", str)
  rq_modcd(0) = Trim(USRDEPT)
  Set rs_modcd = rq_modcd.OpenResultset(1, 1)
  cbo_modality.clear
  cbo_modality.AddItem "ALL"
  While rs_modcd.EOF = False
    cbo_modality.AddItem rs_modcd(0)
    rs_modcd.MoveNext
  Wend
  USRDEPT = "ALL"
    'STR = "select dept_code, dept_name from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ') "
    
    If Trim(USRDEPT) = "ALL" Then
        str = "SELECT * FROM PATOOLSLIB.MA106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='IZZ' )"
        Set SELMA106_QR_ALL = rdb.CreateQuery("selma106_qr_ALL", str)
        'SELMA106_QR(0) = Trim(USRDEPT)
        Set SELMA106_RS_ALL = SELMA106_QR_ALL.OpenResultset(2, 3)
        SELMA106_RS_ALL.MoveFirst
        While Not SELMA106_RS_ALL.EOF
            cmb_dept.AddItem Trim(SELMA106_RS_ALL("dept_name"))
            cmb_dept.Text = Trim(SELMA106_RS_ALL("dept_name"))
            SELMA106_RS_ALL.MoveNext
        Wend
    Else
    
     If Trim(UID) = "A101370" Then
         str = "select dept_code, dept_name from PATOOLSLIB.ma106 WHERE DEPT_CODE=? OR DEPT_CODE='UZZ'"
        Else
             
               str = "select dept_code, dept_name from PATOOLSLIB.ma106 WHERE DEPT_CODE=? "
           
    End If
    
    
       
        Set SELMA106_QR = rdb.CreateQuery("selma106_qr1", str)
        SELMA106_QR(0) = Trim(USRDEPT)
        Set SELMA106_RS = SELMA106_QR.OpenResultset(2, 3)
        SELMA106_RS.MoveFirst
        While Not SELMA106_RS.EOF
            cmb_dept.AddItem Trim(SELMA106_RS("dept_name"))
            cmb_dept.Text = Trim(SELMA106_RS("dept_name"))
            SELMA106_RS.MoveNext
        Wend
    End If
'**************SECURITY**************************************

'If Trim(CMB_DEPT.Text) = "PATHOLOGY" Then
'    CHK_EDIT.Visible = True
'Else
'    CHK_EDIT.Visible = False
'End If
  
  str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USR_ID=?"
  Set SELRISLNKUSR_QR = rdb.CreateQuery("SELRISLNKUSR1", str)
  
  SELRISLNKUSR_QR(0) = "A" & Trim(usr_name)
  Set SELRISLNKUSR_RS = SELRISLNKUSR_QR.OpenResultset(2, 3)
  '101370
  'STR = "SELECT distinct(FREPORTBY) FROM RISLIB.RREQUISITN"
  If Trim(UID) = "A101370" Then
  str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE IN(9,8,6,5) AND (DEPT_CODE=? OR DEPT_CODE='UZZ')"
  Else
    If USRDEPT = "" Then
    str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE IN(9,8,6,5) AND (DEPT_CODE=? OR USR_ID='A101370')"
    Else
    str = "SELECT USERNAME FROM DISLIB.DISLNKUSR WHERE USRTYPE IN(9,8,6,5) AND (DEPT_CODE=?)"
    End If
  End If
  
  Set SELREQN_QR = rdb.CreateQuery("SELREQN_QR061", str)
  'SELREQN_QR(0) = 6
  SELREQN_QR(0) = Trim(USRDEPT)
  Set SELREQN_RS = SELREQN_QR.OpenResultset(2, 3)
  If SELREQN_RS.RowCount > 0 Then
    SELREQN_RS.MoveFirst
    cmb_doctor.clear
    If Trim(USRGRP) = "HOD" Or Trim(USRGRP) = "SYSADM" Then
        While Not SELREQN_RS.EOF
            If Len(Trim(SELREQN_RS("USERNAME"))) > 1 Then
                cmb_doctor.AddItem Trim(SELREQN_RS("USERNAME"))
            End If
            SELREQN_RS.MoveNext
        Wend
    Else
        cmb_doctor.Visible = False
        txt_doctor.Text = Trim(SELRISLNKUSR_RS("USERNAME"))
        txt_doctor.Enabled = True
    End If
End If

'**************SECURITY**************************************
msf_reqstat.Rows = 1
 
 str = "select caseno,reqnno,apntdt,FREPORTDT,EXAMDT from Rislib.Rrequisitn where PREPORTDT between ? and ? and modalitycd=? and freportby=? AND REQNSTAT='T' ORDER BY REQNNO,CASENO"
 Set selreqnn_qr = rdb.CreateQuery("selreqnn_qr1", str)
 selreqnn_qr(0) = "01/01/2000"
 selreqnn_qr(1) = "01/01/2000"
 selreqnn_qr(2) = "CT"
 selreqnn_qr(3) = "XXX"
 Set selreqnn_rs = selreqnn_qr.OpenResultset(1, 3)
 
 ''FOR ALL MODALITY BETWEEN DATES
 str = "select caseno,reqnno,apntdt,FREPORTDT,EXAMDT from Rislib.Rrequisitn where PREPORTDT between ? and ? and freportby=? AND REQNSTAT='T' ORDER BY REQNNO,CASENO"
 Set selreqnn_qr1 = rdb.CreateQuery("selreqnn_qr190", str)
 selreqnn_qr1(0) = "01/01/2000"
 selreqnn_qr1(1) = "01/01/2000"
 selreqnn_qr1(2) = "XXX"
 Set selreqnn_rs1 = selreqnn_qr1.OpenResultset(1, 3)
 
 
 str = "select caseno,reqnno,APNTDT,EXAMDT,FREPORTDT from Rislib.rrequisitn where  modalitycd=? and freportby=? AND REQNSTAT='T' ORDER BY REQNNO,CASENO "
 Set SELREQUI_QR = rdb.CreateQuery("SELREQUI_QR11", str)
 SELREQUI_QR(0) = "CT"
 SELREQUI_QR(1) = "XXX"
 Set SELREQUI_RS = SELREQUI_QR.OpenResultset(1, 3)
 
 ''********FOR ALL MODALITY**************
 str = "select caseno,reqnno,apntdt,EXAMDT,FREPORTDT from Rislib.rrequisitn where  freportby=? AND REQNSTAT='T' ORDER BY REQNNO,CASENO "
 Set SELALL_QR = rdb.CreateQuery("SELALL_QR11", str)
 SELALL_QR(0) = "XXX"
 Set SELALL_RS = SELALL_QR.OpenResultset(1, 3)
 ''********FOR ALL MODALITY**************
 
' STR = "select * from rislib.risauth where usrgroup=? and scrn_desc=?"
'Set RISAU21_QR = rdb.CreateQuery("risau21_qr11", STR)
'RISAU21_QR(0) = Trim(USRGRP)
'RISAU21_QR(1) = "DIS WORKLIST FOR FINAL REPORT"
'Set RISAU21_RS = RISAU21_QR.OpenResultset(1, 1)
'If RISAU21_RS.RowCount = 0 Then
'    MsgBox "No Button Authorisation for this User Group & Screen"
'     End
'End If
CMDPRINT
cmdview
End Sub

Private Sub Form_Unload(cancel As Integer)
Dim QR As rdoQuery
Dim RS As rdoResultset

For Each RS In rdb.rdoResultsets
 RS.Close
Next

For Each QR In rdb.rdoQueries
 QR.Close
Next

End Sub

Private Sub QRCLOSE()
Dim QR As rdoQuery
Dim RS As rdoResultset

For Each RS In rdb.rdoResultsets
 RS.Close
Next

For Each QR In rdb.rdoQueries
 QR.Close
Next

End Sub
Private Sub MaskEd_from_date_LostFocus()
'MaskEd_to_date.SetFocus
End Sub

Private Sub MaskEd_to_date_LostFocus()

Dim i As Integer
Screen.MousePointer = 11
If Opt_case.Value = True And (Trim(cmb_doctor.Text) <> "" Or Trim(txt_doctor.Text) <> "") Then
    Opt_case.Enabled = False
    msf_reqstat.FormatString = "SR.NO. |^   PATH NO.   |^         CASE NO. |^   REQUISITION NO.       |^ SAMPLE REC. DATE |^ COMMIT  |^ REPORT DATE"
    i = 1
    msf_reqstat.Rows = 1
    If Trim(cbo_modality.Text) <> "ALL" Then
        selreqnn_qr(0) = Format(MaskEd_from_date.Text, "yyyy-mm-dd")
        selreqnn_qr(1) = Format(MaskEd_to_date.Text, "yyyy-mm-dd")
        selreqnn_qr(2) = Trim(cbo_modality.Text)
        If Trim(USRGRP) = "HOD" Or Trim(USRGRP) = "SYSADM" Then
            selreqnn_qr(3) = Trim(cmb_doctor.Text)
        Else
            selreqnn_qr(3) = Trim(txt_doctor.Text)
        End If
        selreqnn_rs.MoreResults
        selreqnn_rs.Requery
        If selreqnn_rs.RowCount > 0 Then
            selreqnn_rs.MoveFirst
            While Not selreqnn_rs.EOF
                If SELPATH_FLAG = False Then
                    SELPATH_QR(0) = Trim(selreqnn_rs("REQNNO"))
                    Set SELPATH_RS = SELPATH_QR.OpenResultset(1, 3)
                    SELPATH_FLAG = True
                Else
                    
                    SELPATH_QR(0) = Trim(selreqnn_rs("REQNNO"))
                    Set SELPATH_RS = SELPATH_QR.OpenResultset(1, 3)
                    SELPATH_RS.MoreResults
                    SELPATH_RS.Requery
                End If
                
                If Mid(Trim(selreqnn_rs("reqnno")), 4, 2) = "HL" Then
                
                
                 Dim rs_PBSREPT As rdoResultset
                    Set rs_PBSREPT = rdb.OpenResultset("SELECT * FROM DISLIB.DMULPROCE WHERE REQNO='" & selreqnn_rs("reqnno") & "' AND (PRO_CODE='U713' OR PRO_CODE='U753' OR PRO_CODE='U754' OR PRO_CODE='U709') ", 1, 1)
                    If Not rs_PBSREPT.EOF Then
                    msf_reqstat.AddItem i & vbTab & SELPATH_RS("PATHNO") & vbTab & selreqnn_rs("caseno") & vbTab & req_chglen(selreqnn_rs("reqnno")) & vbTab & selreqnn_rs("APNTDT") & vbTab & vbTab & selreqnn_rs("FREPORTDT")
                    End If
                
                
                Else
                msf_reqstat.AddItem i & vbTab & SELPATH_RS("PATHNO") & vbTab & selreqnn_rs("caseno") & vbTab & req_chglen(selreqnn_rs("reqnno")) & vbTab & selreqnn_rs("APNTDT") & vbTab & vbTab & selreqnn_rs("FREPORTDT")
                End If
                selreqnn_rs.MoveNext
                i = i + 1
            Wend
        Else
            MsgBox "NO PROVISIONAL REPORTS ARE THERE"
        End If
    Else
        selreqnn_qr1(0) = Format(MaskEd_from_date.Text, "yyyy-mm-dd")
        selreqnn_qr1(1) = Format(MaskEd_to_date.Text, "yyyy-mm-dd")
        If Trim(USRGRP) = "HOD" Or Trim(USRGRP) = "SYSADM" Then
            selreqnn_qr1(2) = Trim(cmb_doctor.Text)
        Else
            selreqnn_qr1(2) = Trim(txt_doctor.Text)
        End If
        selreqnn_rs1.MoreResults
        selreqnn_rs1.Requery
        If selreqnn_rs1.RowCount > 0 Then
            selreqnn_rs1.MoveFirst
            While Not selreqnn_rs1.EOF
                If SELPATH_FLAG = False Then
                    SELPATH_QR(0) = Trim(selreqnn_rs1("REQNNO"))
                    Set SELPATH_RS = SELPATH_QR.OpenResultset(1, 3)
                    'SELPATH_FLAG = True
                Else
                    SELPATH_QR(0) = Trim(selreqnn_rs1("REQNNO"))
                    SELPATH_RS.MoreResults
                    SELPATH_RS.Requery
                End If
                
                
                If Mid(Trim(selreqnn_rs1("reqnno")), 3, 2) = "HL" Then
                Else
                End If
                
                msf_reqstat.AddItem i & vbTab & SELPATH_RS("PATHNO") & vbTab & selreqnn_rs1("caseno") & vbTab & req_chglen(selreqnn_rs1("reqnno")) & vbTab & selreqnn_rs1("APNTDT") & vbTab & vbTab & selreqnn_rs1("FREPORTDT")
                selreqnn_rs1.MoveNext
                i = i + 1
            Wend
        Else
            MsgBox "NO PROVISIONAL REPORTS ARE THERE"
        End If
    End If
End If
Screen.MousePointer = 0
'CMD_CANCEL.SetFocus
MaskEd_to_date.Enabled = False
End Sub

Private Sub msf_reqstat_Click()
'On Error GoTo Errhandler:
Dim str1 As String
Dim str2 As String
Dim STR3 As String

DoEvents
Me.MousePointer = vbHourglass
msf_reqstat.col = 2
str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
VAR_REQNNO = req_chglen(msf_reqstat.TextMatrix(msf_reqstat.row, 3))  '''req_chglen(mask_edit_req.Text)
var_caseno = msf_reqstat.TextMatrix(msf_reqstat.row, 2)

If Opt_Pat.Value = False Then
    
    If Len(Trim(var_caseno)) > 6 Then
       Opt_Pat.Value = True
    Else
       Opt_Pat.Value = False
    End If
    
    OPT_CASENO.Value = False
    OPT_CASENO.Enabled = False
    OPT_CASENO.Visible = False
    Label1.Visible = False
    cmb_reqno.Visible = False
    Msk_CCNo.Visible = False
    
End If


SELDREUISITION_QR(0) = VAR_REQNNO
Set SELDREUISITION_RS = SELDREUISITION_QR.OpenResultset(1, 3)


SELRREUISITION_QR(0) = VAR_REQNNO
Set SELRREUISITION_RS = SELRREUISITION_QR.OpenResultset(1, 3)

For i = 0 To 33
    CrystalReport1.Formulas(i) = ""
Next i
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"   'AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'If SELTA101_FLAG = False Then
'    SELTA101_QR(0) = SELRREUISITION_RS("CASENO")
'    Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
'    SELTA101_FLAG = True
'Else
'    SELTA101_QR(0) = SELRREUISITION_RS("CASENO")
'    SELTA101_RS.MoreResults
'    SELTA101_RS.Requery
'End If
'If SELEMPBASINFOFLAG = False Then
'    SELEMPBASINFO_QR(0) = Trim(var_caseno)
'    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
'    SELEMPBASINFOFLAG = True
'Else
'    SELEMPBASINFO_QR(0) = Trim(var_caseno)
'    SELEMPBASINFO_RS.MoreResults
'    SELEMPBASINFO_RS.Requery
'End If


selta1012_qr(0) = SELRREUISITION_RS("CASENO")
Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1

'*** SURGICAL PAHOLOGY REPORT

'*** IMMUNIOHISTO CHEMISTRY REPORT
''If SELRREUISITION_RS("MODALITYCD") = "IH" Then
''    If Len(Trim(var_caseno)) > 6 Then
''        If Len(Trim(SELDREUISITION_RS("SPIHREQNO"))) = 0 Then
''            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISIH.RPT"
''            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
''            If SELRREUISITION_RS.RowCount > 0 Then
''               If Trim(SELRREUISITION_RS("REQNSTAT")) = "T" Then
''                    CrystalReport1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHC REPORT" & "'"
''                Else
''                    CrystalReport1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHC REPORT" & "'"
''                End If
''            End If
''            If selta1012_rs.RowCount > 0 Then
''                CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
''            Else
''                CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
''            End If
''            If SELDREUISITION_RS.RowCount > 0 Then
''                CrystalReport1.Formulas(3) = "PREPATH = '" & Trim(SELDREUISITION_RS("prepathno")) & "'"
''            Else
''                CrystalReport1.Formulas(3) = "PREPATH = '" & "" & "'"
''            End If
''            CrystalReport1.Formulas(4) = "USRID = '" & STR & "'"
''            CrystalReport1.WindowState = crptMaximized
''    '        CRYSTALREPORT1.WindowState = crptMaximized
''            CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
''
''            CrystalReport1.Action = 1
''        End If
''    Else
''        If Len(Trim(SELDREUISITION_RS("SPIHREQNO"))) = 0 Then
''            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISIHSTAFF.RPT"
''            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
''            If SELRREUISITION_RS.RowCount > 0 Then
''               If Trim(SELRREUISITION_RS("REQNSTAT")) = "T" Then
''                    CrystalReport1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHC REPORT" & "'"
''                Else
''                    CrystalReport1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHC REPORT" & "'"
''                End If
''            End If
''            If selta1012_rs.RowCount > 0 Then
''                CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
''            Else
''                CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
''            End If
''            If SELDREUISITION_RS.RowCount > 0 Then
''                CrystalReport1.Formulas(3) = "PREPATH = '" & Trim(SELDREUISITION_RS("prepathno")) & "'"
''            Else
''                CrystalReport1.Formulas(3) = "PREPATH = '" & "" & "'"
''            End If
''            CrystalReport1.Formulas(4) = "USRID = '" & STR & "'"
''            CrystalReport1.Formulas(5) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
''            CrystalReport1.Formulas(6) = "SEX = '" & SELEMPBASINFO_RS("GENDER") & "'"
''            CrystalReport1.Formulas(4) = "SIT = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT"))
''            CrystalReport1.WindowState = crptMaximized
''    '       CRYSTALREPORT1.WindowState = crptMaximized
''            CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
''
''            CrystalReport1.Action = 1
''        End If
''    End If
''  ' *** IMMUNIO HISTO CHEMISTRY REPORT
''End If
 
''If SELRREUISITION_RS("MODALITYCD") = "SP" Then
''  ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
''    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
''    If Len(Trim(var_caseno)) > 6 Then
''
'''        If SELRREUISITION_RS("COMMITFLAG") = "Y" Then
'''           SELDADDENDUMQR(0) = VAR_REQNNO
'''           Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
'''            If SELDADDENDUMRS.RowCount > 0 Then
'''                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
'''            Else
'''                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
'''            End If
'''        Else
''            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
'''        End If
''        If SELDREUISITION_RS.RowCount > 0 Then
''             CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(SELDREUISITION_RS("prepathno")) & "'"
''        Else
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
''        End If
''        'selta1012_qr(0) = selrrequisitnrs("CASENO")
''        'Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
''        If selta1012_rs.RowCount > 0 Then
''            CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
''        Else
''            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
''        End If
''
''        Dim Z
''        Dim Y
''        SELMA301_QR(0) = SELDREUISITION_RS("SITE")
''        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
''        If SELMA301_RS.RowCount > 0 Then
''            CrystalReport1.Formulas(3) = "SIT = '" & Trim(SELMA301_RS("SITEDESC")) & "'"
''        Else
''            CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
''        End If
''        If SELDREUISITION_RS.RowCount > 0 Then
''            If Trim(SELDREUISITION_RS("OPPROCEDUR")) <> "" Then
''                Z = InStr(1, Trim(SELDREUISITION_RS("OPPROCEDUR")), "'")
''                If Z > 0 Then
''                    Y = Trim(SELDREUISITION_RS("OPPROCEDUR"))
''                    U = Len(Y)
''                    X = Left(Y, Z - 1)
''                    w = Right(Y, U - (Z))
''                    Y = X & w
''                    CrystalReport1.Formulas(4) = "OPPRO = '" & Y & "'"
''                Else
''                    CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(SELDREUISITION_RS("OPPROCEDUR")) & "'"
''                End If
''            Else
''                CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(SELDREUISITION_RS("OPPROCEDUR")) & "'"
''            End If
''        Else
''            CrystalReport1.Formulas(4) = "OPPRO = '" & "" & "'"
''        End If
''        STR1 = ""
''        If Trim(SELRREUISITION_RS("REQNSTAT")) = "T" Then
''            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
''            STR1 = "Entered By " & Trim(SELRREUISITION_RS("enterby"))
''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
''        End If
''        SELDFREPORT_QR(0) = VAR_REQNNO
''        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
''        If SELDFREPORT_RS.RowCount > 0 Then
''            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
''                CrystalReport1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
''            Else
''                CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
''            End If
''            If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
''                CrystalReport1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
''            Else
''                CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
''            End If
''            If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
''                CrystalReport1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
''            Else
''                CrystalReport1.Formulas(10) = "GROSDESC = '" & "" & "'"
''            End If
''            If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
''                CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
''            Else
''                CrystalReport1.Formulas(11) = "MICRODESC = '" & "" & "'"
''            End If
''            If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
''                CrystalReport1.Formulas(12) = "IMPN = '" & "Impression" & "'"
''            Else
''                CrystalReport1.Formulas(12) = "IMPN = '" & "" & "'"
''            End If
''
''        End If
''
''        CrystalReport1.Formulas(8) = "USRID = '" & STR & "'"
''
''        'CrystalReport1.Formulas(10) = "GROSDESC = '" & "Gross description" & "'"
''        'CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
''        CrystalReport1.WindowState = crptMaximized
'''        CRYSTALREPORT1.WindowState = crptMaximized
''        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
''
''        CrystalReport1.Action = 1
''    Else
'''        If SELRREUISITION_RS("COMMITFLAG") = "Y" Then
'''            SELDADDENDUMQR(0) = VAR_REQNNO
'''            Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
'''            If SELDADDENDUMRS.RowCount > 0 Then
'''                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADDSTAFF.RPT"
'''            Else
'''                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
'''            End If
'''        Else
''        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
'''        End If
''        SELDREUISITION_RS(0) = VAR_REQNNO
''        Set SELDREUISITION_RS = SELDREUISITION_QR.OpenResultset(1, 3)
''        If SELDREUISITION_RS.RowCount > 0 Then
''             CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(SELDREUISITION_RS("prepathno")) & "'"
''        Else
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
''        End If
''        selta1012_qr(0) = SELRREUISITION_RS("CASENO")
''        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
''        If selta1012_rs.RowCount > 0 Then
''            CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
''        Else
''            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
''        End If
''
''
''        SELMA301_QR(0) = SELDREUISITION_RS("SITE")
''        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
''        If SELMA301_RS.RowCount > 0 Then
''            CrystalReport1.Formulas(3) = "SIT = '" & Trim(SELMA301_RS("SITEDESC")) & "'"
''        Else
''            CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
''        End If
''        If SELDREUISITION_RS.RowCount > 0 Then
''            CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(SELDREUISITION_RS("OPPROCEDUR")) & "'"
''        Else
''            CrystalReport1.Formulas(4) = "OPPRO = '" & "" & "'"
''        End If
''        STR1 = ""
''        If Trim(SELRREUISITION_RS("REQNSTAT")) = "T" Then
''            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
''            STR1 = "Entered By " & Trim(SELRREUISITION_RS("enterby"))
''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
''        Else
''            CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
''            CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
''        End If
''        SELDFREPORT_QR(0) = VAR_REQNNO
''        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
''        If SELDFREPORT_RS.RowCount > 0 Then
''            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
''                CrystalReport1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
''            Else
''                CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
''            End If
''            If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
''                CrystalReport1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
''            Else
''                CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
''            End If
''            If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
''                CrystalReport1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
''            Else
''                CrystalReport1.Formulas(10) = "GROSDESC = '" & "" & "'"
''            End If
''            If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
''                CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
''            Else
''                CrystalReport1.Formulas(11) = "MICRODESC = '" & "" & "'"
''            End If
''            If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
''                CrystalReport1.Formulas(12) = "IMPN = '" & "Impression" & "'"
''            Else
''                CrystalReport1.Formulas(12) = "IMPN = '" & "" & "'"
''            End If
''
''        End If
''
''        CrystalReport1.Formulas(8) = "USRID = '" & STR & "'"
''        CrystalReport1.Formulas(13) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
''        CrystalReport1.Formulas(14) = "DOB = '" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT")) & "'"
''        CrystalReport1.Formulas(15) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
''        CrystalReport1.WindowState = crptMaximized
'''        CRYSTALREPORT1.WindowState = crptMaximized
''        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
''        CrystalReport1.Action = 1
''    End If
'' ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
'' End If
 '**** FNAC REPORT
'''If SELRREUISITION_RS("MODALITYCD") = "FA" Then
'''     If Len(Trim(var_caseno)) > 6 Then
'''        If SELRREUISITION_RS("REQNSTAT") <> "F" Then
'''            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFA.RPT"
'''            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'''            If Trim(SELRREUISITION_RS("REQNSTAT")) = "T" Then
'''                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
'''                STR1 = ""
'''                STR1 = "Entered By " & SELRREUISITION_RS("ENTERBY")
'''                CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'''             Else
'''                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
'''                CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'''            End If
'''            CrystalReport1.Formulas(7) = "USRID = '" & STR & "'"
'''        Else
'''            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADD.RPT"
'''            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'''            CrystalReport1.Formulas(7) = "USRID = '" & STR & "'"
'''        End If
'''        CrystalReport1.WindowState = crptMaximized
''''        CRYSTALREPORT1.WindowState = crptMaximized
'''        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'''
'''        CrystalReport1.Action = 1
'''    Else
'''        If SELRREUISITION_RS("REQNSTAT") <> "F" Then
'''            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFASTAFF.RPT"
'''            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'''            If Trim(SELRREUISITION_RS("REQNSTAT")) = "T" Then
'''                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
'''                STR1 = ""
'''                STR1 = "Entered By " & SELRREUISITION_RS("ENTERBY")
'''                CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'''            Else
'''                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
'''                CrystalReport1.Formulas(9) = "ENT = '" & STR1 & "'"
'''            End If
'''            CrystalReport1.Formulas(7) = "USRID = '" & STR & "'"
'''        Else
'''            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADDSTAFF.RPT"
'''            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'''            CrystalReport1.Formulas(7) = "USRID = '" & STR & "'"
'''        End If
'''        CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'''        CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
'''        CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT"))
'''        CrystalReport1.WindowState = crptMaximized
''''       CRYSTALREPORT1.WindowState = crptMaximized
'''        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'''
'''        CrystalReport1.Action = 1
'''    End If
'''End If
 '**** FNAC REPORT
 '**** CYTOLOGY REPORT
If SELRREUISITION_RS("MODALITYCD") = "CG" Or SELRREUISITION_RS("MODALITYCD") = "CN" Then
     If Len(Trim(var_caseno)) > 6 Then
        If SELRREUISITION_RS("REQNSTAT") <> "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCG.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            If Trim(TXT_CSITE.Text) = "" Then
'                CrystalReport1.Formulas(1) = "STL = '" & "" & "'"
'            Else
'                CrystalReport1.Formulas(1) = "STL = '" & "Site  " & TXT_CSITE.Text & "'"
'            End If
            If Trim(SELRREUISITION_RS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & Trim(SELRREUISITION_RS("ENTERBY"))
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
            CrystalReport1.Formulas(12) = "NOSLIDES=" & "" & SELDREUISITION_RS("LABNO")
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADD.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
            CrystalReport1.Formulas(12) = "NOSLIDES = " & "" & SELDREUISITION_RS("LABNO")
        End If
        If Trim(SELRREUISITION_RS("PREPORTBY")) <> "" And Trim(SELRREUISITION_RS("FREPORTBY")) <> "" And UCase(Trim(SELRREUISITION_RS("PREPORTBY"))) <> UCase(Trim(SELRREUISITION_RS("FREPORTBY"))) Then
            CrystalReport1.Formulas(11) = "P= '" & Trim(SELRREUISITION_RS("PREPORTBY")) & "'"
            CrystalReport1.Formulas(12) = "F= '" & Trim(SELRREUISITION_RS("FREPORTBY")) & "'"
            CrystalReport1.Formulas(13) = "PCAP= '" & "Cytotechnologist" & " '"
            CrystalReport1.Formulas(14) = "FCAP= '" & "Pathologist" & " '"
        End If
        If Trim(SELRREUISITION_RS("PREPORTBY")) <> "" And Trim(SELRREUISITION_RS("FREPORTBY")) = "" Then
            CrystalReport1.Formulas(11) = "PONLY= '" & Trim(SELRREUISITION_RS("PREPORTBY")) & "'"
            'CrystalReport1.Formulas(12) = "F= '" & "" & "'"
            CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
        End If
        If Trim(SELRREUISITION_RS("PREPORTBY")) <> "" And Trim(SELRREUISITION_RS("FREPORTBY")) <> "" And UCase(Trim(SELRREUISITION_RS("PREPORTBY"))) = UCase(Trim(SELRREUISITION_RS("FREPORTBY"))) Then
            CrystalReport1.Formulas(11) = "PONLY= '" & Trim(SELRREUISITION_RS("PREPORTBY")) & "'"
            CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
        End If
    Else
        If SELRREUISITION_RS("REQNSTAT") <> "F" Then
            'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGSTAFF.RPT"
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGS.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            If Trim(TXT_CSITE.Text) = "" Then
'                CrystalReport1.Formulas(1) = "STL = '" & "" & "'"
'            Else
'                CrystalReport1.Formulas(1) = "STL = '" & "Site  " & TXT_CSITE.Text & "'"
'            End If
            If Trim(SELRREUISITION_RS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL CYTOLOGY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & Trim(SELRREUISITION_RS("ENTERBY"))
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL CYTOLOGY REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        Else
            'CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADDSTAFF.RPT"
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADDS.RPT"
            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL CYTOLOGY REPORT WITH ADDENDUM" & "'"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        End If
        CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT"))
    End If
    CrystalReport1.WindowState = crptMaximized
'   CRYSTALREPORT1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
 End If
 '**** CYTOLOGY REPORT
' reports for all modalities
 If SELRREUISITION_RS("MODALITYCD") = "SS" Then
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
    If (IsNull(SELDTMTAB_RS("ABO")) Or Trim(SELDTMTAB_RS("ABO")) = "") And (IsNull(SELDTMTAB_RS("RH")) Or Trim(SELDTMTAB_RS("RH")) = "") And (IsNull(Trim(SELDTMTAB_RS("ABS"))) Or Trim(SELDTMTAB_RS("ABS")) = "") And (IsNull(Trim(SELDTMTAB_RS("IAT"))) Or Trim(SELDTMTAB_RS("IAT")) = "") And (IsNull(Trim(SELDTMTAB_RS("DAT"))) Or Trim(SELDTMTAB_RS("DAT")) = "") And (IsNull(Trim(SELDTMTAB_RS("HBSAG"))) Or Trim(SELDTMTAB_RS("HBSAG")) = "") And (IsNull(Trim(SELDTMTAB_RS("HCV"))) Or Trim(SELDTMTAB_RS("HCV")) = "") And (IsNull(Trim(SELDTMTAB_RS("HIV"))) Or Trim(SELDTMTAB_RS("HIV")) = "") Then
    'If (IsNull(SELDTMTAB_RS("ABS")) Or Trim(SELDTMTAB_RS("ABS")) = "") And (IsNull(Trim(SELDTMTAB_RS("IAT"))) Or Trim(SELDTMTAB_RS("IAT")) = "") Then
        MsgBox " Provisional has not Created !!!", vbInformation, "DTM"
        Exit Sub
    Else
        GoTo DTM
    End If
End If
'''DTM BLANK REPORT CHEQUE
DTM:
'On Error Resume Next
If LNK_FLAG = False Then
    LNK_QR(0) = LOGIN
    Set LNK_RS = LNK_QR.OpenResultset(1, 3)
    LNK_FLAG = True
    Else
    Set LNK_RS = LNK_QR.OpenResultset(1, 3)
    LNK_QR(0) = LOGIN
    LNK_RS.MoreResults
    LNK_RS.Requery
End If

'If selrrequisitnrs("MODALITYCD") = "SP" Then
'If Mid(msf_reqstat.TextMatrix(msf_reqstat.Row, 7), 1, 3) = "PRO" Then
 '      MsgBox "Provisional Report Cannot be viewed"
 '     Exit Sub
'End If
'End If

If SELTA101_FLAG = False Then
    SELTA101_QR(0) = SELRREUISITION_RS("CASENO")
    Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
    SELTA101_FLAG = True
Else
    SELTA101_QR(0) = SELRREUISITION_RS("CASENO")
    'CHANGED BY SHWETA 03-06-2014
    Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
    ''''
    SELTA101_RS.MoreResults
    SELTA101_RS.Requery
End If
Dim NAME1 As String
Dim AGE11
Dim SEX1 As String
If InStr(SELRREUISITION_RS("CASENO"), "/") Then
NAME1 = Trim(SELTA101_RS("TAG")) & " " & Trim(SELTA101_RS("FIRST_NAME")) & " " & Trim(SELTA101_RS("LAST_NAME"))
AGE11 = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), SELRREUISITION_RS("REQNDT"))
SEX1 = Trim(SELTA101_RS("SEX"))
Else

End If

If SELEMPBASINFOFLAG = False Then
    SELEMPBASINFO_QR(0) = SELRREUISITION_RS("caseno")
    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
    SELEMPBASINFOFLAG = True
Else
    SELEMPBASINFO_QR(0) = SELRREUISITION_RS("caseno")
    SELEMPBASINFO_RS.MoreResults
    SELEMPBASINFO_RS.Requery
End If

selta1012_qr(0) = SELRREUISITION_RS("CASENO")
Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"



selta1012_qr(0) = SELRREUISITION_RS("CASENO")
Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)

If SELRREUISITION_RS("MODALITYCD") = "BF" Then
    BFREPORT
End If
If SELRREUISITION_RS("MODALITYCD") = "HL" Then
   VAR_REQNNO = req_chglen(Left(msf_reqstat.Text, 16))
    hlreport1
End If
If SELRREUISITION_RS("MODALITYCD") = "EK" Then
   EKREPORT
End If

'''********BONE MARROW*************
For VAR11 = 0 To 30 '..................................
CrystalReport1.Formulas(VAR11) = ""
Next VAR11

If SELRREUISITION_RS("MODALITYCD") = "BM" Then
    If SELDBMAREPORT_FLAG = False Then
        SELDBMAREPORT_QR(0) = VAR_REQNNO
        Set SELDBMAREPORT_RS = SELDBMAREPORT_QR.OpenResultset(1, 3)
        SELDBMAREPORT_FLAG = True
    Else
        SELDBMAREPORT_QR(0) = VAR_REQNNO
        SELDBMAREPORT_RS.MoreResults
        SELDBMAREPORT_RS.Requery
    End If
   Dim varcaseno As String
   varcaseno = msf_reqstat.TextMatrix(msf_reqstat.row, 2)
    If Opt_Pat.Value = True Then
       SELDBMASUB_QR(0) = VAR_REQNNO
       Set SELDBMASSUB_RS = SELDBMASUB_QR.OpenResultset(1, 1)
       If SELDBMASSUB_RS.RowCount > 0 Then
          If SELDBMASSUB_RS("FULLREP") = "Y" Then
          
        ' CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISBM1.RPT"
          CrystalReport1.ReportFileName = "c:\DISREPORTSEClf\disbm3a.rpt"
          'CrystalReport1.SelectionFormula = "{repbmlf_1.REQnNO} = '" & VAR_REQNNO & "'" ' AND {repbmlf_1.REQNSTAT} <> '' AND {repbmlf_1.REQNSTAT} <> 'M' "
          CrystalReport1.SelectionFormula = "{repbmlf_1.caseno} = '" & varcaseno & "' and {repbmlf_1.REQnNO} = '" & VAR_REQNNO & "'" ' AND {repbmlf_1.REQNSTAT} <> '' AND {repbmlf_1.REQNSTAT} <> 'M' "
          CrystalReport1.Formulas(10) = "NAM='" & Trim(NAME1) & "'"
          CrystalReport1.Formulas(11) = "STDOB='" & Trim(SEX1) & "'"
          CrystalReport1.Formulas(30) = "DOB1=" & Trim(AGE11) & ""
          
          End If
          
         If SELDBMASSUB_RS("FULLREP") = "N" Then
          
        ' CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISBM1.RPT"
          CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBM3AI.RPT"
         End If
       End If
       Dim RS_NOCOMM As rdoResultset
        Set RS_NOCOMM = rdb.OpenResultset("SELECT * FROM DISLIB.repbmlf_1 WHERE REQnNO = '" & Trim(VAR_REQNNO) & "'")
       
        If RS_NOCOMM("COMMENTS") <> "" Or IsNull(RS_NOCOMM("COMMENTS")) = False Then
            CrystalReport1.Formulas(26) = "nocomments = '" & "" & "'"
            CrystalReport1.Formulas(28) = "comments = '" & "COMMENTS" & "'"
        Else
            CrystalReport1.Formulas(26) = "nocomments = 'YES'"
            CrystalReport1.Formulas(28) = "comments = '" & "" & "'"
        End If
       
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Not selta1012_rs.EOF Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & selta1012_rs("PRVDIAGDES") & "'"
        End If
        
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & SELRREUISITION_RS("ENTERBY")
            If Len(Trim(SELRREUISITION_RS("ENTERBY"))) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
            End If
         
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
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
    'CRYSTALREPORT1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=TMH;PWD=TMH;DSN=DISDSN"
       'CrystalReport1.
    CrystalReport1.Action = 1
End If


''******BLOOD PRODUCT ****************

If SELRREUISITION_RS("MODALITYCD") = "BP" Then
    If SELDTMNO_FLAG = False Then
        SELDTMNO_QR(0) = SELRREUISITION_RS("CASENO")
        Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
        SELDTMNO_FLAG = True
    Else
        SELDTMNO_QR(0) = SELRREUISITION_RS("CASENO")
        SELDTMNO_RS.MoreResults
        SELDTMNO_RS.Requery
    End If
    If Opt_Pat.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBP.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(1) = "MOD = '" & "BLOOD PRODUCT" & "'"
        If SELRREUISITION_RS("PATNSTAT") = "OP" Then
            CrystalReport1.Formulas(2) = "ST = '" & SELRREUISITION_RS("pat_catg") & "/" & " Out Patient" & "'"
        Else
            CrystalReport1.Formulas(2) = "ST = '" & SELRREUISITION_RS("pat_catg") & "/" & " In patient" & "'"
        End If
        CrystalReport1.Formulas(3) = "SEXAGE = '" & SELTA101_RS("SEX") & "/" & DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), SELRREUISITION_RS("REQNDT")) & "'"
        If SELRREUISITION_RS("PATNSTAT") = "IP" Then
            CrystalReport1.Formulas(4) = "WRBED = '" & SELRREUISITION_RS("WARD_NAME") & "/" & SELRREUISITION_RS("bedno") & "'"
        Else
            CrystalReport1.Formulas(4) = "WRBED = '" & " " & "'"
        End If
        CrystalReport1.Formulas(5) = "PATHNO = '" & SELDTMNO_RS("DTMNO") & "'"
        CrystalReport1.Formulas(6) = "USRID = '" & str & "'"
    Else
        
    End If
    CrystalReport1.WindowState = crptMaximized
    'CRYSTALREPORT1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport1.Action = 1
End If
''******BLOOD PRODUCT ****************
''******SCREENING TEST ****************
If SELRREUISITION_RS("MODALITYCD") = "SS" Then
    If Opt_Pat.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtm.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If SELRREUISITION_RS.RowCount > 0 Then
            CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(SELRREUISITION_RS("prepathno")) & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
    Else
        CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtmreq.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If SELRREUISITION_RS.RowCount > 0 Then
            CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(SELRREUISITION_RS("prepathno")) & "'"
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
        CrystalReport1.Formulas(5) = "DOB='" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT")) & "'"
    End If
    CrystalReport1.WindowState = crptMaximized
    'CRYSTALREPORT1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport1.Action = 1
End If
''*****SCREENING TEST ****************
''*******HLA TYPING *****************
If SELRREUISITION_RS("MODALITYCD") = "HT" Then
    If Opt_Pat.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISHLA.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        If Len(Trim(SELRREUISITION_RS("LABNO"))) > 0 Then
            CrystalReport1.Formulas(8) = "NAM = '" & SELRREUISITION_RS("NMETERIALS") & "'"
            CrystalReport1.Formulas(9) = "RELATION = 'Relation' "
        Else
            CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELTA101_RS("FIRST_NAME")) + " " + Trim(SELTA101_RS("MID_NAME")) + " " + Trim(SELTA101_RS("LAST_NAME")) & "'"
            CrystalReport1.Formulas(9) = "RELATION ='" & " " & "'"
        End If
        CrystalReport1.Formulas(9) = "dtmno='" & " " & "'"
        CrystalReport1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISHLAstaff.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        If Len(Trim(SELRREUISITION_RS("LABNO"))) > 0 Then
            CrystalReport1.Formulas(8) = "NAM = '" & SELRREUISITION_RS("NMETERIALS") & "'"
            CrystalReport1.Formulas(9) = "RELATION = 'Relation' "
        Else
            CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
            CrystalReport1.Formulas(9) = "RELATION ='" & " " & "'"
        End If
        CrystalReport1.Formulas(10) = "STDOB ='" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(11) = "DOB ='" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT")) & "'"
        CrystalReport1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    End If
End If
''*******HLA TYPING *****************
 '*** SURGICAL PAHOLOGY REPORT
 '*** IMMUNIOHISTO CHEMISTRY REPORT
 If SELRREUISITION_RS("MODALITYCD") = "IH" Then
    If Opt_Pat.Value = True Then
        If Len(Trim(SELRREUISITION_RS("SPIHREQNO"))) = 0 Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISIH.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If SELRREUISITION_RS.RowCount > 0 Then
                CrystalReport1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHC REPORT" & "'"
            End If
            If selta1012_rs.RowCount > 0 Then
                CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
            Else
                CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
            End If
            If SELRREUISITION_RS.RowCount > 0 Then
                CrystalReport1.Formulas(3) = "PREPATH = '" & Trim(SELRREUISITION_RS("prepathno")) & "'"
            Else
                CrystalReport1.Formulas(3) = "PREPATH = '" & "" & "'"
            End If
            CrystalReport1.Formulas(4) = "USRID = '" & str & "'"
            CrystalReport1.WindowState = crptMaximized
            'CRYSTALREPORT1.WindowState = crptMaximized
            CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
            CrystalReport1.Action = 1
        End If
    Else
        If Len(Trim(SELRREUISITION_RS("SPIHREQNO"))) = 0 Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISIHSTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If SELRREUISITION_RS.RowCount > 0 Then
                 CrystalReport1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHC REPORT" & "'"
            End If
            If selta1012_rs.RowCount > 0 Then
                CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
            Else
                CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
            End If
            If SELRREUISITION_RS.RowCount > 0 Then
                CrystalReport1.Formulas(3) = "PREPATH = '" & Trim(SELRREUISITION_RS("prepathno")) & "'"
            Else
                CrystalReport1.Formulas(3) = "PREPATH = '" & "" & "'"
            End If
            CrystalReport1.Formulas(4) = "USRID = '" & str & "'"
            CrystalReport1.Formulas(5) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
            CrystalReport1.Formulas(6) = "SEX = '" & SELEMPBASINFO_RS("GENDER") & "'"
            CrystalReport1.Formulas(4) = "SIT = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT"))
            CrystalReport1.WindowState = crptMaximized
            'CRYSTALREPORT1.WindowState = crptMaximized
            CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
            CrystalReport1.Action = 1
        End If
    End If
  ' *** IMMUNIO HISTO CHEMISTRY REPORT
End If
 
If SELRREUISITION_RS("MODALITYCD") = "SP" Then

  ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
    CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
    If Opt_Pat.Value = True Then
        
        If SELRREUISITION_RS("COMMITFLAG") = "Y" Then
           SELDADDENDUMQR(0) = VAR_REQNNO
           Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
            If SELDADDENDUMRS.RowCount > 0 Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL HISTOPATHOLOGY REPORT" & "'"
            End If
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL HISTOPATHOLOGY REPORT" & "'"
        End If
        seldrequisitn_qr(0) = VAR_REQNNO
        Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
        If seldrequisitn_rs.RowCount > 0 Then
             CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        selta1012_qr(0) = seldrequisitn_rs("CASENO")
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        If selta1012_rs.RowCount > 0 Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        Else
            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
        End If
        
        
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
        
            'CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREUISITION_RS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        
           ' CrystalReport1.Formulas(10) = "FN ='FA' "
        
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
        
        'CrystalReport1.Formulas(10) = "GROSDESC = '" & "Gross description" & "'"
        'CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
        CrystalReport1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    Else
        If SELRREUISITION_RS("COMMITFLAG") = "Y" Then
            SELDADDENDUMQR(0) = VAR_REQNNO
            Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
            If SELDADDENDUMRS.RowCount > 0 Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADDSTAFF.RPT"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL HISTOPATHOLOGY REPORT" & "'"
            End If
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL HISTOPATHOLOGY REPORT" & "'"
        End If
        seldrequisitn_qr(0) = VAR_REQNNO
        Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
        If seldrequisitn_rs.RowCount > 0 Then
             CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        End If
        selta1012_qr(0) = SELRREUISITION_RS("CASENO")
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        If selta1012_rs.RowCount > 0 Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        Else
            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
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
       
            'CrystalReport1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREUISITION_RS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        
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
        CrystalReport1.Formulas(13) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(14) = "DOB = '" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT")) & "'"
        CrystalReport1.Formulas(15) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.WindowState = crptMaximized
'        CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    End If
 ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
 End If
 '**** FNAC REPORT
If SELRREUISITION_RS("MODALITYCD") = "FA" Then
    If Opt_Pat.Value = True Then
        If SELRREUISITION_RS("REQNSTAT") <> "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFA.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            
                If Trim(SELRREUISITION_RS("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                    str1 = ""
                    str1 = "Entered By " & SELRREUISITION_RS("ENTERBY")
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                Else
                    str1 = ""
                    str1 = "Entered By " & SELRREUISITION_RS("ENTERBY")
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                End If
                    CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        End If
             CrystalReport1.Formulas(10) = "FN ='FA '"
             CrystalReport1.WindowState = crptMaximized
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
    Else
        If SELRREUISITION_RS("REQNSTAT") <> "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFASTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            
                If Trim(SELRREUISITION_RS("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                    str1 = ""
                    str1 = "Entered By " & SELRREUISITION_RS("ENTERBY")
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                Else
                    str1 = ""
                    str1 = "Entered By " & SELRREUISITION_RS("ENTERBY")
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                End If
             
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        
        End If
      
        CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT"))
      
        CrystalReport1.WindowState = crptMaximized
'        CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

        CrystalReport1.Action = 1
    End If
End If
 '**** FNAC REPORT

 '**** BIOCHEMISTRY REPORT
 If SELRREUISITION_RS("MODALITYCD") = "BC" Then
    If Opt_Pat.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBC.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        
        'CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & SELRREUISITION_RS("enterby")
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
           
'        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
'        If pcount <> 0 Then
'            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
'        Else
        CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
'        End If
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBCSTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
           CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
           str1 = ""
           str1 = "Entered By " & SELRREUISITION_RS("enterby")
           CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        
'        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
'        If pcount <> 0 Then
'            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
'        Else
        CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
'        End If
        
          CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
          CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
          CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT"))
     
    End If
    CrystalReport1.WindowState = crptMaximized
  '  CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
 End If
 '**** BIOCHEMISTRY REPORT
 '**** EMERGENCY LAB REPORT
 If SELRREUISITION_RS("MODALITYCD") = "EL" Then
    If Opt_Pat.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEL.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        If Trim(prepath) = "" Then
            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
        Else
            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & SELRREUISITION_RS("ENTERBY")
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         
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
        
            CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & SELRREUISITION_RS("ENTERBY")
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
         
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        If pcount <> 0 Then
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        End If
      
         CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
         CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
         CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT"))
     
    End If
    CrystalReport1.WindowState = crptMaximized
  '  CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
 End If
 '**** EMERGENCY LAB REPORT
 '**** TUMOUR MARKER REPORT
 If SELRREUISITION_RS("MODALITYCD") = "TM" Then
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
     If Opt_Pat.Value = True Then
         If IMFLAG1 = True Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(prepath) = "" Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
            End If
            
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                str1 = ""
                str1 = "Entered by " & SELRREUISITION_RS("ENTERBY")
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        End If
        If IMFLAG = True Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM1.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(prepath) = "" Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
            End If
            
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                str1 = ""
                str1 = "Entered by " & SELRREUISITION_RS("ENTERBY")
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            
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
            CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                str1 = ""
                str1 = "Entered by " & SELRREUISITION_RS("ENTERBY")
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
            If SELEMPBASINFO_RS.RowCount > 0 Then
                CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
                CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
                CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT"))
            End If
        End If
        If IMFLAG = True Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISTM1STAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If Trim(prepath) = "" Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
            End If
            If Trim(SELRREUISITION_RS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                str1 = ""
                str1 = "Entered by " & SELRREUISITION_RS("ENTERBY")
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
                CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
                CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT"))
      
        End If
    End If
    CrystalReport1.WindowState = crptMaximized
   ' CRYSTALREPORT1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport1.Action = 1
End If
'
 IMFLAG = False
 IMFLAG1 = False
 
 '**** TUMOUR MARKER REPORT
 '****ecg reports
If SELRREUISITION_RS("MODALITYCD") = "PF" Then
    If Opt_Pat.Value = True Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISPF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
    '        If Trim(prepath) = "" Then
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
    '        End If
        selta1012_qr(0) = SELRREUISITION_RS("CASENO")
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREUISITION_RS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISPFSTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
    '        If Trim(prepath) = "" Then
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
    '        End If
        selta1012_qr(0) = SELRREUISITION_RS("CASENO")
        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
        If Not selta1012_rs.EOF Then
            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
        Else
            CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
        End If
        If Trim(SELRREUISITION_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREUISITION_RS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        
        CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT"))
        
    End If
    CrystalReport1.WindowState = crptMaximized
'    CRYSTALREPORT1.WindowState = crptMaximized
    CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
End If
If SELRREUISITION_RS("MODALITYCD") = "ST" Then
    If Opt_Pat.Value = True Then
       CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISST.RPT"
       CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
    '        If Trim(prepath) = "" Then
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
    '        Else
    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
    '        End If
       CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
       
           CrystalReport1.Formulas(5) = "PROFIN = '" & "STRESS TEST REPORT" & "'"
           str1 = ""
           str1 = "Entered By " & Trim(SELRREUISITION_RS("enterby"))
           If Len(Trim(SELRREUISITION_RS("enterby"))) > 0 Then
               CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
           Else
               CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
           End If
           CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
       
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
           CrystalReport1.Formulas(5) = "PROFIN = '" & "STRESS TEST REPORT" & "'"
           str1 = ""
           str1 = "Entered By " & Trim(SELRREUISITION_RS("enterby"))
           If Len(Trim(SELRREUISITION_RS("enterby"))) > 0 Then
               CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
           Else
               CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
           End If
           CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
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
        CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), SELRREUISITION_RS("REQNDT"))
    End If
    CrystalReport1.WindowState = crptMaximized
'    CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
End If
If SELRREUISITION_RS("MODALITYCD") = "FI" Or SELRREUISITION_RS("MODALITYCD") = "CC" Then
      CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCY.RPT"
      CrystalReport1.WindowState = crptMaximized
'        CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

        CrystalReport1.Action = 1
End If


 '****ecg report
 '**** RIS REPORTS
''If selrrequisitnrs("MODALITYCD") = "CT" Or selrrequisitnrs("MODALITYCD") = "US" Or selrrequisitnrs("MODALITYCD") = "MM" Or selrrequisitnrs("MODALITYCD") = "MR" Or selrrequisitnrs("MODALITYCD") = "CR" Then
''    rq_grid_det(0) = Trim(VAR_REQNNO)
''    rs_grid_det.MoreResults
''    rs_grid_det.Requery
''    If rs_grid_det.RowCount > 0 And (Trim(rs_grid_det("REQNSTAT"))) = "T" Then
''        If Opt_patient.Value = True Then
''            CrystalReport1.Formulas(0) = "provtyp= '" & Trim(rs_grid_det("ENTERBY")) & "'"
''            'cr_rep.Formulas(1) = "reporttext= '" & Trim(txt_report.Text) & "'"
''            'cr_rep.Formulas(0) = "impresstext= '" & Trim(txt_impression.Text) & "'"
''            'cr_rep.Formulas(1) = "finalrep= '" & Trim(txt_finalreport.Text) & "'"
''
''            SELRMODALITY_QR(0) = Mid(VAR_REQNNO, 4, 2)
''            SELRMODALITY_RS.MoreResults
''            SELRMODALITY_RS.Requery
''            If SELRMODALITY_RS.RowCount > 0 Then
''                CrystalReport1.Formulas(1) = "MODDES='" & SELRMODALITY_RS("MODLITYDES") & "'"
''            Else
''                CrystalReport1.Formulas(1) = "MODDES='" & "" & "'"
''            End If
''            CrystalReport1.Formulas(2) = "deptname= '" & "Department of Radiodiagnosis" & "'"
''            CrystalReport1.SelectionFormula = "{RISREPORTS.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
''            'cr_rep.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
''            ''cr_rep.Formulas(0) = "provtypist = '" & txt_type1.Text & "'"
''            STR1 = ip
''            STR2 = DATE11
''            STR3 = time1
''            STR1 = STR1 & Space(5) & STR2 & Space(5) & STR3
''            STR1 = Space(3) & STR1
''            CrystalReport1.Formulas(3) = "us_id = '" & STR1 & "'"
''            CrystalReport1.Formulas(4) = "INOUT = '" & Trim(rs_grid_det("PATNSTAT")) & "'"
''            CrystalReport1.Formulas(5) = "reqno = '" & Left(msf_reqstat.Text, 16) & "'"
''            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep1.rpt"
''
''        Else
''            CrystalReport1.Formulas(0) = "provtyp= '" & Trim(rs_grid_det("ENTERBY")) & "'"
''            'cr_rep.Formulas(1) = "reporttext= '" & Trim(txt_report.Text) & "'"
''            'cr_rep.Formulas(0) = "impresstext= '" & Trim(txt_impression.Text) & "'"
''            'cr_rep.Formulas(1) = "finalrep= '" & Trim(txt_finalreport.Text) & "'"
''
''            SELRMODALITY_QR(0) = Mid(VAR_REQNNO, 4, 2)
''            SELRMODALITY_RS.MoreResults
''            SELRMODALITY_RS.Requery
''            If SELRMODALITY_RS.RowCount > 0 Then
''                CrystalReport1.Formulas(1) = "MODDES='" & SELRMODALITY_RS("MODLITYDES") & "'"
''            Else
''                CrystalReport1.Formulas(1) = "MODDES='" & "" & "'"
''            End If
''            CrystalReport1.Formulas(2) = "deptname= '" & "Department of Radiodiagnosis" & "'"
''            CrystalReport1.SelectionFormula = "{RISREPORTS.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
''            'cr_rep.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
''            ''cr_rep.Formulas(0) = "provtypist = '" & txt_type1.Text & "'"
''            STR1 = ip
''            STR2 = DATE11
''            STR3 = time1
''            STR1 = STR1 & Space(5) & STR2 & Space(5) & STR3
''            STR1 = Space(3) & STR1
''            CrystalReport1.Formulas(3) = "us_id = '" & STR1 & "'"
''            CrystalReport1.Formulas(4) = "INOUT = '" & Trim(rs_grid_det("PATNSTAT")) & "'"
''            CrystalReport1.Formulas(5) = "reqno = '" & Left(msf_reqstat.Text, 16) & "'"
''            CrystalReport1.Formulas(6) = "PATNAME = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
''            CrystalReport1.Formulas(7) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
''            CrystalReport1.Formulas(5) = "PATAGE = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
''            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\crpt_rep1STAFF.rpt"
''        End If
''        CrystalReport1.WindowState = crptMaximized
'''        CRYSTALREPORT1.WindowState = crptMaximized
''        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
''
''        CrystalReport1.Action = 1
''        Exit Sub
''    Else
''        Exit Sub
''    End If
''End If
 '**** RIS REPORTS
 '*************************************************************************
 '          wrote for microbiology ****************************************
 ' on 12/12/2002
 If SELRREUISITION_RS("MODALITYCD") = "BA" Or SELRREUISITION_RS("MODALITYCD") = "MB" Or SELRREUISITION_RS("MODALITYCD") = "XX" Or SELRREUISITION_RS("MODALITYCD") = "SE" Or SELRREUISITION_RS("MODALITYCD") = "CM" Then
    printed = "True"
    Microbiology SELRREUISITION_RS("MODALITYCD")
 ElseIf SELRREUISITION_RS("MODALITYCD") = "2D" Then
    printed = "True"
    Echocardiology SELRREUISITION_RS("MODALITYCD")
 ElseIf SELRREUISITION_RS("MODALITYCD") = "ED" Then
    printed = "True"
    Endoscopy SELRREUISITION_RS("MODALITYCD")
 End If

' reports for all modalities
 
 If CrystalReport1.RecordsPrinted = 1 Or printed = "True" Then
     msf_reqstat.col = 5
     If Trim(msf_reqstat.TextMatrix(msf_reqstat.row, 5)) = "Y" Then
       chk_Commit.Value = 1
     Else
       chk_Commit.Value = 0
     End If
     chk_Commit.Left = msf_reqstat.CellLeft + msf_reqstat.Left
     chk_Commit.Top = msf_reqstat.CellTop + msf_reqstat.Top
     chk_Commit.Width = msf_reqstat.CellWidth
     chk_Commit.Visible = True
         
     msf_reqstat.col = 7
     If Trim(msf_reqstat.TextMatrix(msf_reqstat.row, 7)) = "Y" Then
       chk_edit.Value = 1
     Else
       chk_edit.Value = 0
     End If
     chk_edit.Left = msf_reqstat.CellLeft + msf_reqstat.Left
     chk_edit.Top = msf_reqstat.CellTop + msf_reqstat.Top
     chk_edit.Width = msf_reqstat.CellWidth
     chk_edit.Visible = True
    
 Else
     chk_Commit.Value = False
     chk_edit.Value = False
 End If
Me.MousePointer = vbDefault
Exit Sub
Errhandler:
MsgBox Err.Description, vbCritical, "WORK LIST"
Exit Sub
End Sub

Private Sub Msk_CCNo_LostFocus()
    If Msk_CCNo.Text = "" Or Msk_CCNo.Text = "__/______" Or Msk_CCNo.Text = "______" Then
        MsgBox ("Case No. Cannot be left Blank"), vbInformation
        Exit Sub
    End If
    
If Opt_Pat.Value = True Then
    Msk_CCNo.Text = padzero(5, "C", Msk_CCNo.Text, 3)
End If

'freportby=? AND REQNSTAT='T'AND DEPT_CODE='FZZ'"

sel_ReqNo_qr(0) = Trim(Msk_CCNo.FormattedText)
If Trim(USRGRP) = "SYSADM" Or Trim(USRGRP) = "HOD" Or Trim(USRGRP) = "MEDADMN" Then
    sel_ReqNo_qr(1) = Trim(cmb_doctor.Text)
Else
    sel_ReqNo_qr(1) = Trim(txt_doctor.Text)
End If
    If sel_ReqNo_flg = False Then
        Set sel_ReqNo_Rs = sel_ReqNo_qr.OpenResultset(1, 3)
        sel_ReqNo_flg = True
    Else
        sel_ReqNo_Rs.MoreResults
        sel_ReqNo_Rs.Requery
    End If
    cmb_reqno.clear
    
    If sel_ReqNo_Rs.RowCount > 0 Then
        cmb_reqno.AddItem "ALL"
        sel_ReqNo_Rs.MoveFirst
        While Not sel_ReqNo_Rs.EOF
            cmb_reqno.AddItem Trim(sel_ReqNo_Rs("reqnno"))
            sel_ReqNo_Rs.MoveNext
        Wend
     '   sel_ReqNo_Rs.Close
        
   End If
End Sub

Private Sub Opt_all_LostFocus()
'msf_reqstat.Refresh
Dim i As String
Screen.MousePointer = 11
If OPT_ALL.Value = True And (Trim(cmb_doctor.Text) <> "" Or Trim(txt_doctor.Text) <> "") Then
    Opt_case.Enabled = False
    msf_reqstat.FormatString = "SR.NO. |^   PATH NO.   |^           CASE NO. |^   REQUISITION NO.       |^ SAMPLE REC. DATE |^ COMMIT |^  REPORT DATE"
    i = 1
    msf_reqstat.Rows = 1
    If Trim(cbo_modality.Text) = "ALL" Then
        If Trim(USRGRP) = "SYSADM" Or Trim(USRGRP) = "HOD" Or Trim(USRGRP) = "MEDADMN" Then
            SELALL_QR(0) = Trim(cmb_doctor.Text)
        Else
            SELALL_QR(0) = Trim(txt_doctor.Text)
        End If
        SELALL_RS.MoreResults
        SELALL_RS.Requery
        If SELALL_RS.RowCount > 0 Then
            SELALL_RS.MoveFirst
            While Not SELALL_RS.EOF
                If SELPATH_FLAG = False Then
                    SELPATH_QR(0) = SELALL_RS("REQNNO")
                    Set SELPATH_RS = SELPATH_QR.OpenResultset(1, 3)
                    SELPATH_FLAG = True
                Else
                    SELPATH_QR(0) = SELALL_RS("REQNNO")
                    SELPATH_RS.MoreResults
                    SELPATH_RS.Requery
                End If
                msf_reqstat.AddItem i & vbTab & Trim(SELPATH_RS("PATHNO")) & vbTab & Trim(SELALL_RS("caseno")) & vbTab & req_chglen(SELALL_RS("reqnno")) & vbTab & Trim(SELALL_RS("APNTDT")) & vbTab & vbTab & Trim(SELALL_RS("freportdt"))
                SELALL_RS.MoveNext
                i = i + 1
            Wend
        Else
            MsgBox "NO PROVISIONAL REPORTS ARE THERE"
        End If
    Else
        SELREQUI_QR(0) = Trim(cbo_modality.Text)
        If Trim(USRGRP) = "SYSADM" Or Trim(USRGRP) = "HOD" Or Trim(USRGRP) = "MEDADMN" Then
            SELREQUI_QR(1) = Trim(cmb_doctor.Text)
        Else
            SELREQUI_QR(1) = Trim(txt_doctor.Text)
        End If
        SELREQUI_RS.MoreResults
        SELREQUI_RS.Requery
        If SELREQUI_RS.RowCount > 0 Then
            SELREQUI_RS.MoveFirst
            While Not SELREQUI_RS.EOF
                If SELPATH_FLAG = False Then
                    SELPATH_QR(0) = SELREQUI_RS("REQNNO")
                    Set SELPATH_RS = SELPATH_QR.OpenResultset(1, 3)
                    SELPATH_FLAG = True
                Else
                    SELPATH_QR(0) = SELREQUI_RS("REQNNO")
                    SELPATH_RS.MoreResults
                    SELPATH_RS.Requery
                End If
                msf_reqstat.AddItem i & vbTab & SELPATH_RS("PATHNO") & vbTab & SELREQUI_RS("caseno") & vbTab & req_chglen(SELREQUI_RS("reqnno")) & vbTab & SELREQUI_RS("APNTDT") & vbTab & vbTab & SELREQUI_RS("FREPORTDT")
                SELREQUI_RS.MoveNext
                i = i + 1
            Wend
        Else
            MsgBox "NO PROVISIONAL REPORTS ARE THERE !!", vbInformation, "WORK LIST"
            Exit Sub
        End If
    End If
End If
Screen.MousePointer = 0
'CMD_CANCEL.SetFocus
End Sub

Private Sub Opt_case_Click()
If Opt_case.Value = True Then
MaskEd_from_date.Enabled = True
MaskEd_to_date.Enabled = True
MaskEd_from_date.Text = Format(DATE11, "dd/mm/yyyy")
MaskEd_to_date.Text = Format(DATE11, "dd/mm/yyyy")
MaskEd_from_date.SetFocus
OPT_ALL.Enabled = False
Msk_CCNo.Visible = False
Label1.Visible = False
cmb_reqno.Visible = False
OPT_CASENO.Enabled = False

FROMFLAG = True
End If
End Sub

Public Sub cmdview()
'If RISAU21_RS("VIEW") = 1 Then
'CMD_VIEW.Enabled = True
'Else
'CMD_VIEW.Enabled = False
'End If
End Sub
Public Sub CMDPRINT()

'If RISAU21_RS("PRINT") = 1 Then
'CMD_PRINT.Enabled = True
'Else
'CMD_PRINT.Enabled = False
'End If
End Sub

Private Sub Option1_Click()

End Sub



Private Sub Opt_Pat_Click()
    OPT_CASENO.Caption = "Case No."
    OPT_CASENO.Value = True
    OPT_CASENO.Enabled = True
    OPT_CASENO.Visible = True
    Msk_CCNo.Mask = ">??/#####"
    Label1.Visible = True
    cmb_reqno.Visible = True
    Msk_CCNo.Visible = True
    'Msk_CCNo.SetFocus
    MaskEd_to_date.Enabled = False
End Sub

Private Sub opt_staff_Click()
    OPT_CASENO.Caption = "CC No."
    OPT_CASENO.Enabled = True
    OPT_CASENO.Value = True
    OPT_CASENO.Visible = True
    Msk_CCNo.Mask = ">#######"
    Label1.Visible = True
    cmb_reqno.Visible = True
    Msk_CCNo.Visible = True
    MaskEd_to_date.Enabled = False
    Msk_CCNo.SetFocus
End Sub

Private Sub Microbiology(moda As String)
On Error GoTo MSG
Load Report_microbiology
Report_microbiology.Hide
Report_microbiology.show
If moda = "BA" Then
    Report_microbiology.cmb_modality.Text = "BACTERIOLOGY$BA"
    Report_microbiology.printpage "BA", var_caseno, VAR_REQNNO
ElseIf moda = "MB" Then
    Report_microbiology.cmb_modality.Text = "MYCOBACTERIOLOGY$MB"
    Report_microbiology.printpage "MB", var_caseno, VAR_REQNNO
ElseIf moda = "CM" Then
    Report_microbiology.cmb_modality.Text = "CLINICAL MICROBIOLOGY$CM"
    Report_microbiology.printpage "CM", var_caseno, VAR_REQNNO
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

Private Sub BFREPORT()
For i = 0 To 25
    CrystalReport1.Formulas(i) = ""
Next i
'VAR_REQNNO = req_chglen(Left(msf_reqstat.Text, 16))
selta1012_qr(0) = msf_reqstat.TextMatrix(msf_reqstat.row, 2)
Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)

selrrequisitnqr(0) = VAR_REQNNO
Set selrrequisitnrs = selrrequisitnqr.OpenResultset(1, 3)
If Opt_Pat.Value = True Then
             
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"

    
    
     If Trim(selrrequisitnrs("MODALITYCD")) = "BF" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBFold.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
If selta1012_rs.EOF = True Then
  CrystalReport1.Formulas(2) = "PROVDIA = ''"
Else
  CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
End If
      
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
      CrystalReport1.Formulas(13) = "mm = '" & "abc" & "'"
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
VAR_REQNNO = Replace(msf_reqstat.TextMatrix(msf_reqstat.row, 3), "/", "") '''req_chglen(mask_edit_req.Text)
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


'            If SELRREQUISITNFLAG1 = False Then
'                SELRREQUISITN_QR(0) = Trim(MaskEdBox1.FormattedText)
'                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
'                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
'                SELRREQUISITN_QR(3) = "T"
'                SELRREQUISITN_QR(4) = "F"
'                Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
'                SELRREQUISITNFLAG1 = True
'            Else
'                SELRREQUISITN_QR(0) = Trim(MaskEdBox1.FormattedText)
'                SELRREQUISITN_QR(1) = Trim(txt_modcode.Text)
'                SELRREQUISITN_QR(2) = Trim(txt_deptcode.Text)
'                SELRREQUISITN_QR(3) = "T"
'                SELRREQUISITN_QR(4) = "F"
'                SELRREQUISITN_RS.MoreResults
'                SELRREQUISITN_RS.Requery
'            End If
'            If SELRREQUISITN_Rs.RowCount > 0 Then
'                prepath = ""
'                While Not SELRREQUISITN_Rs.EOF
'                    If Trim(SELRREQUISITN_Rs("reqnno")) <> req_chglen(Trim(cmb_reqno.Text)) Then
'                        prepath = prepath & req_chglen(SELRREQUISITN_Rs("REQNNO")) & " "
'                    End If
'
'                    SELRREQUISITN_Rs.MoveNext
'                Wend
'             End If
'
If Opt_Pat.Value = True Then

If SELDHEMCYTO_RS.EOF = True Then
Dim str123 As String

Dim RS_DMULPROCE As rdoResultset
  Set RS_DMULPROCE = rdb.OpenResultset("SELECT * FROM diSLIB.dmulproce WHERE REQNO = '" & Trim(VAR_REQNNO) & "'")
  If RS_DMULPROCE.EOF = False Then
    If RS_DMULPROCE("PRO_CODE") = "U722" Then
    str123 = "insert into DISLIB.dhemcyto1(REQNO,CBC) VALUES ('" & Trim(VAR_REQNNO) & "','YES')"

    ElseIf RS_DMULPROCE("PRO_CODE") = "U713" Then
    str123 = "insert into DISLIB.dhemcyto1(REQNO,MPO,HEM) VALUES ('" & Trim(VAR_REQNNO) & "','YES','YES')"

    ElseIf RS_DMULPROCE("PRO_CODE") = "U706" Then
    str123 = "insert into DISLIB.dhemcyto1(REQNO,ESR) VALUES ('" & Trim(VAR_REQNNO) & "','YES')"
    
    ElseIf RS_DMULPROCE("PRO_CODE") = "U709" Or RS_DMULPROCE("PRO_CODE") = "U712" Or RS_DMULPROCE("PRO_CODE") = "U714" Or RS_DMULPROCE("PRO_CODE") = "U715" Or RS_DMULPROCE("PRO_CODE") = "U708" Or RS_DMULPROCE("PRO_CODE") = "U710" Then
    str123 = "insert into DISLIB.dhemcyto1(REQNO,CG) VALUES ('" & Trim(VAR_REQNNO) & "','YES')"
    
    ElseIf RS_DMULPROCE("PRO_CODE") = "U752" Then
    str123 = "insert into DISLIB.dhemcyto1(REQNO,CYTO) VALUES ('" & Trim(VAR_REQNNO) & "','YES')"
    
     ElseIf RS_DMULPROCE("PRO_CODE") = "U754" Or RS_DMULPROCE("PRO_CODE") = "U753" Then
    str123 = "insert into DISLIB.dhemcyto1(REQNO,SMAR) VALUES ('" & Trim(VAR_REQNNO) & "','YES')"
     ElseIf RS_DMULPROCE("PRO_CODE") = "U755" Then
    str123 = "insert into DISLIB.dhemcyto1(REQNO,HEM,CG) VALUES ('" & Trim(VAR_REQNNO) & "','YES','YES')"
Else
    str123 = "insert into DISLIB.dhemcyto1(REQNO) VALUES ('" & Trim(VAR_REQNNO) & "')"
    End If
  End If

rdb.Execute str123
End If

SELDHEMCYTO_QR(0) = VAR_REQNNO
Set SELDHEMCYTO_RS = SELDHEMCYTO_QR.OpenResultset(1, 1)


  If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("HEM")) = "YES" Or Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
   
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"

    
    
  '  If Trim(txt_modcode.Text) = "HL" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShlN1AA.RPT"
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
           
'          If Trim(txt_lapscore.Text) = "" And Trim(TXT_MPO.Text) = "" And Trim(TXT_SBB.Text) = "" And Trim(TXT_PAS.Text) = "" And Trim(TXT_NSE.Text) = "" And Trim(TXT_TRAP.Text) = "" Then
'           'If CYTOCHEMFLAG = True Then
'             CRYSTALREPORT1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
'           Else
'            CRYSTALREPORT1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
'           'End If
'           End If
           ' If Trim(TXT_BT.Text) = "" And Trim(TXT_PT(0).Text) = "" And Trim(TXT_PLATELETSCP.Text) = "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_TT.Text) = "" And Trim(TXT_FIBRINOGEN.Text) = "" And Trim(TXT_FDPDMIER.Text) = "" And Trim(txt_others.Text) = "" Then

            ' CRYSTALREPORT1.Formulas(11) = "CP = '" & "" & "'"
           'Else
           ' CRYSTALREPORT1.Formulas(11) = "CP = '" & "ABC" & "'"

          ' End If

'           If Trim(TXT_NORMOBLASTS.Text) = "" And Trim(TXT_OTHERCELLSRBC.Text) = "" And Trim(txt_rbcmorphology.Text) = "" And Trim(txt_reticcount.Text) = "" And Trim(txt_malarialparasite.Text) = "" And Trim(TXT_ESRRBC.Text) = "" Then
'
'             CRYSTALREPORT1.Formulas(15) = "RBC = '" & "" & "'"
'           Else
'            CRYSTALREPORT1.Formulas(15) = "RBC = '" & "ABC" & "'"
'
'           End If
'           If Trim(TXT_MORPHOLOGY.Text) = "" And Trim(TXT_ADEQUACY.Text) = "" Then
'
'             CRYSTALREPORT1.Formulas(16) = "PLATELETS = '" & "" & "'"
'           Else
'            CRYSTALREPORT1.Formulas(16) = "PLATELETS = '" & "ABC" & "'"
'
'           End If
'           If Trim(txt_prolymphocytes.Text) = "" And Trim(txt_Sezarycells.Text) = "" And Trim(txt_hairycells.Text) = "" And Trim(txt_othercells.Text) = "" Then
'                CRYSTALREPORT1.Formulas(17) = "prolyn = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(17) = "prolyn = '" & "123" & "'"
'           End If
'           If Trim(txt_blasts.Text) = "" And Trim(TXT_PROMYELOCYTESWBC.Text) = "" And Trim(txt_myelocytes.Text) = "" And Trim(txt_metamyelocytes.Text) = "" And Trim(txt_plasmacells.Text) = "" And Trim(txt_alymphocytes.Text) = "" Then
'                CRYSTALREPORT1.Formulas(18) = "blast = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(18) = "blast = '" & "123" & "'"
'           End If
'           If Trim(txt_mneutrophils.Text) = "" And Trim(txt_mlymphocytes.Text) = "" And Trim(txt_mmonocytes.Text) = "" And Trim(txt_meosinophils.Text) = "" And Trim(txt_mbasophills.Text) = "" Then
'                CRYSTALREPORT1.Formulas(19) = "cbcm = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(19) = "cbcm = '" & "123" & "'"
'           End If
'           If Trim(TXT_HCT.Text) = "" And Trim(TXT_MCH.Text) = "" And Trim(TXT_MCV.Text) = "" And Trim(TXT_MCHC.Text) = "" And Trim(TXT_MPV.Text) = "" And Trim(TXT_RDW.Text) = "" And Trim(TXT_ABNORMAL.Text) = "" Then
'                CRYSTALREPORT1.Formulas(20) = "rbci = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(20) = "rbci = '" & "123" & "'"
'           End If
'           If Trim(txt_neutrophils.Text) = "" And Trim(txt_lymphocytes.Text) = "" And Trim(txt_monocytes.Text) = "" And Trim(txt_eosinophils.Text) = "" And Trim(txt_basophils.Text) = "" And Trim(txt_immaturegranulocytes.Text) = "" Then
'                CRYSTALREPORT1.Formulas(21) = "dlc = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(21) = "dlc = '" & "123" & "'"
'           End If
'           If Trim(txt_hb.Text) = "" And Trim(TXT_RBC.Text) = "" And Trim(TXT_TWBC.Text) = "" And Trim(txt_platelets.Text) = "" Then
'                CRYSTALREPORT1.Formulas(22) = "cbcc = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(22) = "cbcc = '" & "123" & "'"
'           End If
           '' PRINT REPORT
           CrystalReport1.WindowState = crptMaximized
           CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
           CrystalReport1.Action = 1
'           If REP1 = True And REP2 = True Then
'              MsgBox "There are two reports print this one first"
'           End If
           
   '  End If
 End If  'CBC/HEM/CG
 '''  NEW
 If Trim(SELDHEMCYTO_RS("CYTO")) = "YES" Or Trim(SELDHEMCYTO_RS("SMAR")) = "YES" Then
   
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"

    
    
'     If Trim(txt_modcode.Text) = "HL" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShlN1A.RPT"
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
           
'          If Trim(txt_lapscore.Text) = "" And Trim(TXT_MPO.Text) = "" And Trim(TXT_SBB.Text) = "" And Trim(TXT_PAS.Text) = "" And Trim(TXT_NSE.Text) = "" And Trim(TXT_TRAP.Text) = "" Then
'           'If CYTOCHEMFLAG = True Then
'             CRYSTALREPORT1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
'           Else
'            CRYSTALREPORT1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
'           'End If
'           End If
           ' If Trim(TXT_BT.Text) = "" And Trim(TXT_PT(0).Text) = "" And Trim(TXT_PLATELETSCP.Text) = "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_TT.Text) = "" And Trim(TXT_FIBRINOGEN.Text) = "" And Trim(TXT_FDPDMIER.Text) = "" And Trim(txt_others.Text) = "" Then

            ' CRYSTALREPORT1.Formulas(11) = "CP = '" & "" & "'"
          ' Else
          '  CRYSTALREPORT1.Formulas(11) = "CP = '" & "ABC" & "'"

          ' End If

'           If Trim(TXT_NORMOBLASTS.Text) = "" And Trim(TXT_OTHERCELLSRBC.Text) = "" And Trim(txt_rbcmorphology.Text) = "" And Trim(txt_reticcount.Text) = "" And Trim(txt_malarialparasite.Text) = "" And Trim(TXT_ESRRBC.Text) = "" Then
'
'             CRYSTALREPORT1.Formulas(15) = "RBC = '" & "" & "'"
'           Else
'            CRYSTALREPORT1.Formulas(15) = "RBC = '" & "ABC" & "'"
'
'           End If
'           If Trim(TXT_MORPHOLOGY.Text) = "" And Trim(TXT_ADEQUACY.Text) = "" Then
'
'             CRYSTALREPORT1.Formulas(16) = "PLATELETS = '" & "" & "'"
'           Else
'            CRYSTALREPORT1.Formulas(16) = "PLATELETS = '" & "ABC" & "'"
'
'           End If
'           If Trim(txt_prolymphocytes.Text) = "" And Trim(txt_Sezarycells.Text) = "" And Trim(txt_hairycells.Text) = "" And Trim(txt_othercells.Text) = "" Then
'                CRYSTALREPORT1.Formulas(17) = "prolyn = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(17) = "prolyn = '" & "123" & "'"
'           End If
'           If Trim(txt_blasts.Text) = "" And Trim(TXT_PROMYELOCYTESWBC.Text) = "" And Trim(txt_myelocytes.Text) = "" And Trim(txt_metamyelocytes.Text) = "" And Trim(txt_plasmacells.Text) = "" And Trim(txt_alymphocytes.Text) = "" Then
'                CRYSTALREPORT1.Formulas(18) = "blast = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(18) = "blast = '" & "123" & "'"
'           End If
'           If Trim(txt_mneutrophils.Text) = "" And Trim(txt_mlymphocytes.Text) = "" And Trim(txt_mmonocytes.Text) = "" And Trim(txt_meosinophils.Text) = "" And Trim(txt_mbasophills.Text) = "" Then
'                CRYSTALREPORT1.Formulas(19) = "cbcm = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(19) = "cbcm = '" & "123" & "'"
'           End If
'           If Trim(TXT_HCT.Text) = "" And Trim(TXT_MCH.Text) = "" And Trim(TXT_MCV.Text) = "" And Trim(TXT_MCHC.Text) = "" And Trim(TXT_MPV.Text) = "" And Trim(TXT_RDW.Text) = "" And Trim(TXT_ABNORMAL.Text) = "" Then
'                CRYSTALREPORT1.Formulas(20) = "rbci = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(20) = "rbci = '" & "123" & "'"
'           End If
'           If Trim(txt_neutrophils.Text) = "" And Trim(txt_lymphocytes.Text) = "" And Trim(txt_monocytes.Text) = "" And Trim(txt_eosinophils.Text) = "" And Trim(txt_basophils.Text) = "" And Trim(txt_immaturegranulocytes.Text) = "" Then
'                CRYSTALREPORT1.Formulas(21) = "dlc = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(21) = "dlc = '" & "123" & "'"
'           End If
'           If Trim(txt_hb.Text) = "" And Trim(TXT_RBC.Text) = "" And Trim(TXT_TWBC.Text) = "" And Trim(txt_platelets.Text) = "" Then
'                CRYSTALREPORT1.Formulas(22) = "cbcc = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(22) = "cbcc = '" & "123" & "'"
'           End If
'           If Trim(SELDHEMCYTO_RS("smar")) = "YES" Then
'             CRYSTALREPORT1.Formulas(23) = "mar = '" & "" & "'"
'           Else
'             CRYSTALREPORT1.Formulas(23) = "mar = '" & "123" & "'"
'           End If
'
           ' PRINT REPPORT
           
           CrystalReport1.WindowState = crptMaximized
           CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
           CrystalReport1.Action = 1
    ' End If
 End If  'CYTO/SURFACE MARKER
 For VAR11 = 0 To 30 '..................................
CrystalReport1.Formulas(VAR11) = ""
Next VAR11

Dim RS_DMULPROCE1 As rdoResultset
  Set RS_DMULPROCE1 = rdb.OpenResultset("SELECT * FROM diSLIB.dmulproce WHERE REQNO = '" & Trim(VAR_REQNNO) & "' AND PRO_CODE='U753'")
  If RS_DMULPROCE1.EOF = True Then
 CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\HEMATO.RPT"
CrystalReport1.Formulas(0) = "REQNO='" & Trim(VAR_REQNNO) & "'"
'CrystalReport1.Formulas(17) = "signature='" & Trim(cmb_finalby.Text) & "'"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport1.Action = 1
 End If
 
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

 
     If Trim(txt_modcode.Text) = "EK" Then
            CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISEHLASTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
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
    If Trim(txt_modcode.Text) = "HL" Then
'' NEW CODES FOR STAFF REPORTS
  If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("HEM")) = "YES" Or Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
   
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"

    
    
     If Trim(txt_modcode.Text) = "HL" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShlN1AASTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            If Len(Trim(cmb_enterby.Text)) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                 CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
            End If
         Else

            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
           
        '  If Trim(txt_lapscore.Text) = "" And Trim(TXT_MPO.Text) = "" And Trim(TXT_SBB.Text) = "" And Trim(TXT_PAS.Text) = "" And Trim(TXT_NSE.Text) = "" And Trim(TXT_TRAP.Text) = "" Then
           'If CYTOCHEMFLAG = True Then
             CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "" & "'"
         '  Else
            CrystalReport1.Formulas(10) = "CYTOCHEM = '" & "ABC" & "'"
           'End If
        '   End If
           ' If Trim(TXT_BT.Text) = "" And Trim(TXT_PT(0).Text) = "" And Trim(TXT_PLATELETSCP.Text) = "" And Trim(TXT_PTTK.Text) = "" And Trim(TXT_TT.Text) = "" And Trim(TXT_FIBRINOGEN.Text) = "" And Trim(TXT_FDPDMIER.Text) = "" And Trim(txt_others.Text) = "" Then

           '  CRYSTALREPORT1.Formulas(11) = "CP = '" & "" & "'"
           'Else
           ' CRYSTALREPORT1.Formulas(11) = "CP = '" & "ABC" & "'"

           'End If

         '  If Trim(TXT_NORMOBLASTS.Text) = "" And Trim(TXT_OTHERCELLSRBC.Text) = "" And Trim(txt_rbcmorphology.Text) = "" And Trim(txt_reticcount.Text) = "" And Trim(txt_malarialparasite.Text) = "" And Trim(TXT_ESRRBC.Text) = "" Then

         '   CRYSTALREPORT1.Formulas(15) = "RBC = '" & "" & "'"
          ' Else
          '  CRYSTALREPORT1.Formulas(15) = "RBC = '" & "ABC" & "'"

         '  End If
        '   If Trim(TXT_MORPHOLOGY.Text) = "" And Trim(TXT_ADEQUACY.Text) = "" Then

          '   CRYSTALREPORT1.Formulas(16) = "PLATELETS = '" & "" & "'"
         '  Else
         '   CRYSTALREPORT1.Formulas(16) = "PLATELETS = '" & "ABC" & "'"

          ' End If
'           If Trim(txt_prolymphocytes.Text) = "" And Trim(txt_Sezarycells.Text) = "" And Trim(txt_hairycells.Text) = "" And Trim(txt_othercells.Text) = "" Then
'                CRYSTALREPORT1.Formulas(17) = "prolyn = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(17) = "prolyn = '" & "123" & "'"
'           End If
'           If Trim(txt_blasts.Text) = "" And Trim(TXT_PROMYELOCYTESWBC.Text) = "" And Trim(txt_myelocytes.Text) = "" And Trim(txt_metamyelocytes.Text) = "" And Trim(txt_plasmacells.Text) = "" And Trim(txt_alymphocytes.Text) = "" Then
'                CRYSTALREPORT1.Formulas(18) = "blast = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(18) = "blast = '" & "123" & "'"
'           End If
'           If Trim(txt_mneutrophils.Text) = "" And Trim(txt_mlymphocytes.Text) = "" And Trim(txt_mmonocytes.Text) = "" And Trim(txt_meosinophils.Text) = "" And Trim(txt_mbasophills.Text) = "" Then
'                CRYSTALREPORT1.Formulas(19) = "cbcm = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(19) = "cbcm = '" & "123" & "'"
'           End If
'           If Trim(TXT_HCT.Text) = "" And Trim(TXT_MCH.Text) = "" And Trim(TXT_MCV.Text) = "" And Trim(TXT_MCHC.Text) = "" And Trim(TXT_MPV.Text) = "" And Trim(TXT_RDW.Text) = "" And Trim(TXT_ABNORMAL.Text) = "" Then
'                CRYSTALREPORT1.Formulas(20) = "rbci = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(20) = "rbci = '" & "123" & "'"
'           End If
'           If Trim(txt_neutrophils.Text) = "" And Trim(txt_lymphocytes.Text) = "" And Trim(txt_monocytes.Text) = "" And Trim(txt_eosinophils.Text) = "" And Trim(txt_basophils.Text) = "" And Trim(txt_immaturegranulocytes.Text) = "" Then
'                CRYSTALREPORT1.Formulas(21) = "dlc = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(21) = "dlc = '" & "123" & "'"
'           End If
'           If Trim(txt_hb.Text) = "" And Trim(TXT_RBC.Text) = "" And Trim(TXT_TWBC.Text) = "" And Trim(txt_platelets.Text) = "" Then
'                CRYSTALREPORT1.Formulas(22) = "cbcc = '" & "" & "'"
'           Else
'                CRYSTALREPORT1.Formulas(22) = "cbcc = '" & "123" & "'"
'           End If
           '' PRINT REPORT
           CrystalReport1.WindowState = crptMaximized
           CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
           CrystalReport1.Action = 1
'           If REP1 = True And REP2 = True Then
'              MsgBox "There are two reports print this one first"
'           End If
     End If
 End If  'CBC/HEM/CG
 '''  NEW
 If Trim(SELDHEMCYTO_RS("CYTO")) = "YES" Or Trim(SELDHEMCYTO_RS("SMAR")) = "YES" Then
   
    str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
    CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNSTAT} <> '' AND {RREQUISITN.REQNSTAT} <> 'M'"

    
    
     If Trim(txt_modcode.Text) = "HL" Then
         CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DIShlN1ASTAFF.RPT"
        CrystalReport1.Formulas(0) = "REQNO = '" & Trim(cmb_reqno.Text) & "'"
'        If Trim(prepath) = "" Then
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
        CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
        If Trim(REFDR_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "HAEMATOLOGY REPORT" & "'"
            str1 = ""
            str1 = "Entered By " & Trim(cmb_enterby.Text)
            If Len(Trim(cmb_enterby.Text)) > 0 Then
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            Else
                 CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
            End If
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
End If


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

Private Sub Timer1_Timer()
txt_time.Text = Format(Time, "HH:MM:SS")
End Sub
