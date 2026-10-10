VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Barcode 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Sample Receiving Centre"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   16485
   LinkTopic       =   "Form2"
   ScaleHeight     =   10980
   ScaleWidth      =   16485
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_pendlist 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pending List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Refresh Grid"
      Top             =   9600
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1080
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_print_report 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Refresh Grid"
      Top             =   9600
      Width           =   1095
   End
   Begin VB.TextBox txt_pathno 
      BackColor       =   &H00C0FFFF&
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
      Left            =   8760
      MaxLength       =   4
      TabIndex        =   18
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
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
      Height          =   1575
      Left            =   2400
      TabIndex        =   8
      Top             =   1320
      Width           =   8415
      Begin VB.TextBox TXTREQNO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3240
         MaxLength       =   11
         TabIndex        =   10
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cmb_location 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Barcode.frx":0000
         Left            =   3240
         List            =   "Barcode.frx":0002
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisition No :-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   960
         TabIndex        =   12
         Top             =   720
         Width           =   1785
      End
      Begin VB.Label lbllocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Received At :- "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   2460
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00004080&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         Height          =   1335
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   8175
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   6135
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   14895
      Begin VB.TextBox TxtUnit 
         BackColor       =   &H00FFFFC0&
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
         Left            =   11520
         MaxLength       =   10
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5895
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   10398
         _Version        =   393216
         Cols            =   15
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   16761024
         ForeColorFixed  =   128
         GridColor       =   0
         AllowUserResizing=   3
         FormatString    =   "S.No |^ Requisition No.    |^Case No.  |^ Name            |^Modality    |^Rec Date|^RecTime    |^       "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame buttons 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
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
      Left            =   2400
      TabIndex        =   2
      Top             =   9240
      Width           =   10695
      Begin VB.CommandButton cmd_smplrcvPRINT 
         BackColor       =   &H00C0FFC0&
         Caption         =   "PRINT RCV LIST"
         Height          =   375
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Print Barcode"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         MouseIcon       =   "Barcode.frx":0004
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print Report"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmd_back 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Back"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         MouseIcon       =   "Barcode.frx":030E
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Back to the Main Screen"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CMD_REFRESH 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Refresh Grid"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00004080&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         Height          =   855
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   10335
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
      Left            =   15240
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
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
      Left            =   15240
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   1320
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
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
      Left            =   1440
      TabIndex        =   22
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
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
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Requisition belongs to Emergency Lab."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11280
      TabIndex        =   29
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Left            =   10920
      TabIndex        =   28
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Urgent Sample"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11280
      TabIndex        =   27
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000013&
      Height          =   255
      Left            =   10920
      TabIndex        =   26
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "AUTONUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   23
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "02/11/21"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbl_req 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " SAMPLE RECEIVING CENTRE "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   4440
      TabIndex        =   16
      Top             =   960
      Width           =   4125
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   14400
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   14400
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
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
      Left            =   1050
      TabIndex        =   13
      Top             =   360
      Width           =   12075
   End
End
Attribute VB_Name = "Barcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pend As Boolean
Dim SELDEPT_QR As rdoQuery
Dim SELDEPT_RS As rdoResultset
Dim GETCASE_QR As rdoQuery
Dim GETCASE_RS As rdoResultset
Dim GETNAME_QR As rdoQuery
Dim GETNAME_RS As rdoResultset
Dim GETNAMESTAFF_QR As rdoQuery
Dim GETNAMESTAFF_RS As rdoResultset
Dim SELREQNSTAT_QR As rdoQuery
Dim SELREQNSTAT_RS As rdoResultset
Dim UPDRTIME_QR As rdoQuery
Dim UPDRREQ_QR As rdoQuery


Dim UPD_TMAUTOPATH_qr As rdoQuery
Dim sel_TMAUTOPATH_qr As rdoQuery
Dim Sel_TMAUTOPATHANY_rs As rdoResultset
Dim UPD_TMAUTOPATHANY_qr As rdoQuery
Dim Sel_TMAUTOPATHANY_QR As rdoQuery

'NEW FOR MICRO
Dim UPD_MICROAUTOPATH_qr As rdoQuery
Dim sel_MICROAUTOPATH_qr As rdoQuery
Dim sel_MICROAUTOPATH_rs As rdoResultset
Dim Sel_MICROTESTCODE_QR As rdoQuery
Dim sel_MICROTESTCODE_rs As rdoResultset

'----- Shraddha Pandit 27/11/08 ------------
Dim SELSEPATHNO_QR As rdoQuery
Dim SELSEPATHNO_RS As rdoResultset

Dim UPDDAUTOPATH_QR As rdoQuery
Dim UPDDAUTOPATH_RS As rdoResultset

Dim INSDAUTOPATH_QR As rdoQuery
Dim INSDAUTOPATH_RS As rdoResultset

Dim updrrequisitn_qr As rdoQuery

Dim SEPATHNO As Variant
Dim dept As String

Dim INSBINTERFACE_QR As rdoQuery
Dim RS_DMULPROCE  As rdoResultset

Dim SEL_TEMPCOLLCTR_QR As rdoQuery
Dim SEL_TEMPCOLLCTR_rs As rdoResultset

Dim UPDTIME_QR As rdoQuery
Dim UPDTIME_RS As rdoResultset
Dim SELTIME_QR As rdoQuery
Dim SELTIME_RS As rdoResultset
Dim SELTIME_FLG As Boolean
Dim SELREQNOQR As rdoQuery
Dim SELREQNORS As rdoResultset
Dim upddrequisitn As rdoQuery

Dim SEL_TEMPCOLL_QR As rdoQuery
Dim SEL_TEMPCOLL_RS As rdoResultset

Dim INS_TEMPCOLLCT As rdoQuery

Dim upd_TEMMPCOLL_QR As rdoQuery 'Shraddha 04/08/2011
Dim upd_TCOLL_Print_QR As rdoQuery

Dim upd_TCOLL_P_QR As rdoQuery

Dim upd_TCOLL_blank_QR As rdoQuery 'Shraddha 19/10/2011
Dim upd_TCOLL_RCVIP_QR As rdoQuery

Dim seldmulproce_qr As rdoQuery
Dim seldmulproce_rs As rdoResultset
Dim SELDBCPROSUB_QR1 As rdoQuery
Dim SELDBCPROSUB_RS1 As rdoResultset
Dim SELSAMPLECODE_QR As rdoQuery
Dim ins_minterface As rdoQuery

Dim SELDTMNO_QR As rdoQuery
Dim SELDTMNO_RS As rdoResultset
'***** Shraddha 24/12/09 *****
Dim SEL_PathDreq_QR As rdoQuery
Dim SEL_PathDreq_RS As rdoResultset

Dim UPD_RREQNSTAT_QR As rdoQuery
Dim UPD_RREQNSTAT_RS As rdoResultset
'*****************************

Dim SEL_TA101_QR As rdoQuery
Dim SEL_TA101_RS As rdoResultset

Dim REPWTRN_QR As rdoQuery
Dim REPWTRNRS As rdoResultset

'***** Shraddha 12/10/09
Dim SELDAUTOPATH_QR As rdoQuery
Dim SELDAUTOPATH_RS As rdoResultset

Dim DPATHYEAR_QR As rdoQuery
Dim DPATHYEAR_RS As rdoResultset

Dim patstat As String

Dim SELMA101_QR As rdoQuery
Dim SELMA101_RS As rdoResultset

Dim SEL_BINTER_QR As rdoQuery
Dim SEL_BINTER_RS As rdoResultset

'***** Shraddha 14/10/09
Dim INSHINTERFACE_QR As rdoQuery

Dim INS_IHINTERFACE_QR As rdoQuery 'Shraddha 18/11/10

Dim INSCINTERFACE_QR As rdoQuery

Dim SELDHEMCYTO_QR As rdoQuery
Dim SELDHEMCYTO_RS As rdoResultset

Dim INS_CINTERFACE_QR As rdoQuery

Dim seldhlprosub_qr As rdoQuery
Dim seldhlprosub_rs As rdoResultset

Dim pathno As Integer

Dim tmpathno As Double

Dim SELDBCPROSUB_QR As rdoQuery
Dim SELDBCPROSUB_RS As rdoResultset

Dim interexpcd_qr As rdoQuery
Dim interexpxd_rs As rdoResultset

Dim INSDINTERFACE_QR As rdoQuery

Dim seldmulproce_qr1 As rdoQuery
Dim seldmulproce_rs1 As rdoResultset

Dim sel_TMTESCODE_qr As rdoQuery
Dim sel_TMTESCODE_rs As rdoResultset

Dim sel_TMCATEGORY_qr As rdoQuery
Dim sel_TMCATEGORY_rs As rdoResultset

Dim sel_TMCATEGORYRF_qr As rdoQuery
Dim sel_TMCATEGORYRF_RS As rdoResultset

Dim sel_TMCATEGORYANY_qr As rdoQuery
Dim sel_TMCATEGORYANY_rs As rdoResultset

Dim UPD_TMCATEGORY_qr As rdoQuery
Dim UPD_TMCATEGORYANY_qr As rdoQuery
Dim UPD_TMCATEGORYRF_qr As rdoQuery

Dim seldelprosub_qr As rdoQuery
Dim seldelprosub_rs As rdoResultset

Dim SAMPLECODE_QR1 As rdoQuery
Dim SAMPLECODE_RS1 As rdoResultset

Dim insdinterface_qr1 As rdoQuery
Dim fecreqno As String

Dim SELDREQUISITNQR As rdoQuery
Dim SELDREQUISITNRS As rdoResultset

Dim SEL_RREQUP_QR As rdoQuery
Dim SEL_RREQUP_RS As rdoResultset

Dim upd_rtime_qr As rdoQuery '06/11/09
Dim UPD_mTIME_QR As rdoQuery '06/11/09
Dim flg As Boolean
'-------------------------------------------

Dim REQNOB As String
Dim REQNOBARCODE As String
Dim REQ111 As String
Dim PAT_NAME As String
Dim COLLDATE As String
Dim COLLTIME As String

Dim MODALITYB As String
Dim DEPTB As String
Dim LEN1 As Integer
Dim str As String
Public Sub MICROPATH()
'Dim MPATHNO As String
'Dim MAPATHNO As String
Dim prefix1 As String, reqT1 As String
prefix1 = ""




Dim SEL_MA101_QR As rdoQuery
Dim SEL_MA101_RS As rdoResultset
DATETIME
Dim GTESTDES As String
str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
Set SEL_MA101_QR = rdb.CreateQuery("", str)

'START HERE
seldmulproce_qr1(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
Set seldmulproce_rs1 = seldmulproce_qr1.OpenResultset(1, 3)
If Not seldmulproce_rs1.EOF Then '1
      
        
'             If Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), 4, 2) = "BA" Then
'                  If Trim(seldmulproce_rs1("PRO_CODE")) = "G260" Then
'                   prefix1 = "AB"
'                   Else
'                   prefix1 = "BB"
'                  End If
'            ElseIf Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), 4, 2) = "MB" Then
'            prefix1 = "MB"
'            ElseIf Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), 4, 2) = "MC" Then
'            prefix1 = "MC"
'            ElseIf Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), 4, 2) = "MM" Then
'            prefix1 = "MM"
'            ElseIf Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), 4, 2) = "CM" Then
'            prefix1 = "CM"
'            Else
'            prefix1 = ""
'            End If
'
Sel_MICROTESTCODE_QR(0) = Trim(seldmulproce_rs1("PRO_CODE"))
Set sel_MICROTESTCODE_rs = Sel_MICROTESTCODE_QR.OpenResultset(1, 3)
        If Not sel_MICROTESTCODE_rs.EOF Then
   ''''''
        If Trim(sel_MICROTESTCODE_rs("PREFIXC")) <> "" Then
        PREFEX1 = Trim(sel_MICROTESTCODE_rs("PREFIXC"))
        End If
   '''''''''''''
        sel_MICROAUTOPATH_qr(0) = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), 4, 2)
  ''''''''''''
        If Trim(sel_MICROTESTCODE_rs("PREFIXC")) = "" Or IsNull(sel_MICROTESTCODE_rs("PREFIXC")) Then
        sel_MICROAUTOPATH_qr(1) = sel_MICROTESTCODE_rs("TYPE")
        Else
        sel_MICROAUTOPATH_qr(1) = Trim(sel_MICROTESTCODE_rs("PREFIXC"))
        End If
        Set sel_MICROAUTOPATH_rs = sel_MICROAUTOPATH_qr.OpenResultset(1, 3)
            If sel_MICROAUTOPATH_rs.EOF = False Then
            l = Len(Trim(sel_MICROAUTOPATH_rs("PATHNO"))) ''VARINDER
            prefix1 = Trim(sel_MICROAUTOPATH_rs("PREFIX1"))
            If Trim(seldmulproce_rs1("PRO_CODE")) = "G202" Then
            prefix1 = "MS"
            End If
            ''''Monika'''
            If Trim(sel_MICROTESTCODE_rs("PREFIXC")) <> "" Then
            prefix1 = Trim(sel_MICROTESTCODE_rs("PREFIXC"))
            End If
            '''''''''''''''''''''''''05/06/2023
                If UCase(Trim(sel_MICROAUTOPATH_rs("STATUS"))) = "MONTH" Or UCase(Trim(sel_MICROAUTOPATH_rs("STATUS"))) = "YEAR1" Then
                '''for month
                'PR = Trim(Mid(DATE11, 4, 2))
                PR = Trim(Format(DATE11, "DDMMMYY"))
'                   If Trim(Mid(DATE11, 4, 2)) = "01" Or Trim(Mid(DATE11, 4, 2)) = "02" Or Trim(Mid(DATE11, 4, 2)) = "03" Or Trim(Mid(DATE11, 4, 2)) = "04" Or Trim(Mid(DATE11, 4, 2)) = "05" Or Trim(Mid(DATE11, 4, 2)) = "06" Or Trim(Mid(DATE11, 4, 2)) = "07" Or Trim(Mid(DATE11, 4, 2)) = "08" Or Trim(Mid(DATE11, 4, 2)) = "09" Then '999
'
'                        If Mid(sel_MICROAUTOPATH_rs("PATHNO"), 3, 1) = Trim(Mid(DATE11, 5, 1)) Then
'                           MPATHNO = Trim(Mid(sel_MICROAUTOPATH_rs("PATHNO"), 7, l - 1))
'                           MAPATHNO = MPATHNO + 1
'                          Else
'                           MPATHNO = 1
'                          ' MAPATHNO = (Mid(DATE11, 4, 2)) & "0001"
'                           MAPATHNO = "0001"
'                          End If
'            '*********** 999 if bet oct - Dec **********************
'                     ElseIf Trim(Mid(DATE11, 4, 2)) = "10" Or Trim(Mid(DATE11, 4, 2)) = "11" Or Trim(Mid(DATE11, 4, 2)) = "12" Then '999
                       If Trim(sel_MICROAUTOPATH_rs("MONTH1")) = Val(Trim(Mid(DATE11, 4, 2))) Then
                           MPATHNO = Trim(sel_MICROAUTOPATH_rs("PATHNO"))  'Added on 02/11/09
                           MAPATHNO = MPATHNO + 1
                         Else
                           MPATHNO = 1
                          ' MAPATHNO = (Mid(DATE11, 4, 2)) & "0001"
                           MAPATHNO = "0001"
                        End If
                    ' End If
                     
                     '''''newly added for prefix
                     
                     
                     
                     MAPATHNO = MAPATHNO
                     
                     '''till here prefix
                 'end here month
                ElseIf UCase(Trim(sel_MICROAUTOPATH_rs("STATUS"))) = "YEAR" Then
                PR = Trim(Format(DATE11, "DDMMMYY"))   'Trim(Mid(DATE11, 9, 2))
                
                  If Trim(Format(DATE11, "YY")) = Trim(sel_MICROAUTOPATH_rs("YEAR1")) Then
                  MPATHNO = Trim(sel_MICROAUTOPATH_rs("PATHNO"))
                           MAPATHNO = MPATHNO + 1
                  Else
                  MPATHNO = 1
                  MAPATHNO = "0001"
                  'MAPATHNO = Trim(Format(DATE11, "DDMMYY")) & "0001"
                  End If
                  
                 ' Dim UPD_MICROAUTOPATH_qr1 As rdoQuery
                 ' str = "UPDATE DISLIB.MICROPATH SET YEAR1=YEAR(CURRENT DATE) ,MONTH1=MONTH(CURRENT DATE) WHERE MODALITY=? AND STATUS=? AND TYPE = ?"
                 ' Set UPD_MICROAUTOPATH_qr1 = rdb.CreateQuery("", str)
                  
                   'If Trim(Mid(DATE11, 9, 2)) = "01" Or Trim(Mid(DATE11, 9, 2)) = "02" Or Trim(Mid(DATE11, 9, 2)) = "03" Or Trim(Mid(DATE11, 9, 2)) = "04" Or Trim(Mid(DATE11, 9, 2)) = "05" Or Trim(Mid(DATE11, 9, 2)) = "06" Or Trim(Mid(DATE11, 9, 2)) = "07" Or Trim(Mid(DATE11, 9, 2)) = "08" Or Trim(Mid(DATE11, 9, 2)) = "09" Then '999
                    '    If Left(sel_MICROAUTOPATH_rs("PATHNO"), 1) = Trim(Mid(DATE11, 10, 1)) Then
                          'MPATHNO = Trim(sel_MICROAUTOPATH_rs("PATHNO"))
                           'MAPATHNO = MPATHNO + 1
                        '  Else
                         '  MPATHNO = 1
                         '  MAPATHNO = Trim(Format(DATE11, "DDMMYY")) & "0001"
                        '  End If
            '*********** 999 if bet oct - Dec **********************
                     'ElseIf Trim(Mid(DATE11, 9, 2)) = "10" Or Trim(Mid(DATE11, 9, 2)) = "11" Or Trim(Mid(DATE11, 9, 2)) = "12" Then '999
                     '  If Left(sel_MICROAUTOPATH_rs("PATHNO"), 2) = Trim(Mid(DATE11, 9, 2)) Then
                        '   MPATHNO = Trim(Mid(sel_MICROAUTOPATH_rs("PATHNO"), 3, l - 1)) 'Added on 02/11/09
                         '  MAPATHNO = MPATHNO + 1
                         'Else
                         '  MPATHNO = 1
                        '   MAPATHNO = Trim(Format(DATE11, "DDMMYY")) & "0001"
                       ' End If
                     'End If
                Else
                   PR = Trim(Format(DATE11, "DDMMYY"))  ' "00"
                    MPATHNO = Trim(Mid(sel_MICROAUTOPATH_rs("PATHNO"), 3, l - 1)) 'Added on 02/11/09 ''VARINDER
                    MAPATHNO = MPATHNO + 1
                End If
                Dim pno As String
                pno = Trim(MAPATHNO)
                    If Len(Trim(MAPATHNO)) = 1 Then
                    MAPATHNO = "000" & MAPATHNO
                    
                    ElseIf Len(Trim(MAPATHNO)) = 2 Then
                    MAPATHNO = "00" & MAPATHNO
                    ElseIf Len(Trim(MAPATHNO)) = 3 Then
                    MAPATHNO = "0" & MAPATHNO
                    End If
                    'MAPATHNO = PR & MAPATHNO
                    MAPATHNO = MAPATHNO
                    
                    UPD_MICROAUTOPATH_qr(0) = MAPATHNO ' (Mid(DATE11, 4, 2)) & SEAUTO
                    UPD_MICROAUTOPATH_qr(1) = Trim(SELDEPT_RS("MODALITYCD"))
                    UPD_MICROAUTOPATH_qr(2) = Trim(sel_MICROAUTOPATH_rs("status"))
                    UPD_MICROAUTOPATH_qr(3) = Trim(sel_MICROAUTOPATH_rs("type"))
                    UPD_MICROAUTOPATH_qr.Execute
                   ' If Trim(seldmulproce_rs1("PRO_CODE")) = "G260" Then
                   ' upddrequisitn(0) = "AB" & MAPATHNO  ' pno & "/" & Format(txt_date.Text, "DDMMYYYY")
                   ' Else
                   
                   
                    upddrequisitn(0) = MAPATHNO & prefix1 & PR
                    'upddrequisitn(0) = pno & "/" & Format(txt_date.Text, "DDMMYYYY")
                 '   End If
                    
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = upddrequisitn(0) 'PNO & "/" & Format(txt_date.Text, "DDMMYYYY")
                   
                    
                      SEL_MA101_QR(0) = Trim(seldmulproce_rs1("PRO_CODE"))
                      Set SEL_MA101_RS = SEL_MA101_QR.OpenResultset(1, 3)
                      GTESTDES = GTESTDES & "|" & Trim(SEL_MA101_RS("SER_DET"))
                    
                    upddrequisitn(1) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
                    upddrequisitn.Execute
            End If 'FOR MICRO AUTO PATH
        End If  'END IF FOR TMTESTCODE
End If 'END FOR DMULPROCE
MSFlexGrid1.TextMatrix(0, 8) = "Test Description"
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 8) = GTESTDES 'Trim(seldmulproce_rs1("PRO_CODE"))

End Sub

Public Sub SEAUTONO()
On Error GoTo MSG:
MODALITY = Trim(SELDEPT_RS("MODALITYCD"))
Dim j
j = 1
For j = 1 To MSFlexGrid1.Rows - 1

    If Trim(MSFlexGrid1.TextMatrix(j, 7)) = "" Then 'Shraddha 15/12/09
    
    If Trim(MSFlexGrid1.TextMatrix(j, 4)) <> "" Then
        'If SELSEPATHNO_FLAG = False Then
            SELSEPATHNO_QR(0) = Trim(dept)
            SELSEPATHNO_QR(1) = Trim(SELDEPT_RS("MODALITYCD"))
            SELSEPATHNO_QR(2) = Trim(Right(DATE11, 2))
            Set SELSEPATHNO_RS = SELSEPATHNO_QR.OpenResultset(1, 3)
            SELSEPATHNO_FLAG = True
'        Else
'            SELSEPATHNO_QR(0) = Trim(DEPT)
'            SELSEPATHNO_QR(1) = Trim(MODALITY)
'            SELSEPATHNO_QR(2) = Trim(Right(DATE11, 2))
'            SELSEPATHNO_RS.MoreResults
'            SELSEPATHNO_RS.Requery
'        End If

'        If SELSEPATHNO_RS.RowCount > 0 Then
'            If Left(SELSEPATHNO_RS("PATHNO"), 2) = Trim(Mid(DATE11, 4, 2)) Then
'                SEAUTO = Mid(SELSEPATHNO_RS("PATHNO"), 3)
'                SEAUTO = SEAUTO + 1
'                UPDDAUTOPATH_QR(0) = (Mid(DATE11, 4, 2)) & SEAUTO
'                UPDDAUTOPATH_QR(1) = Trim(dept)
'                UPDDAUTOPATH_QR(2) = Trim(SELDEPT_RS("MODALITYCD"))
'                UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
'                UPDDAUTOPATH_QR.Execute
'            Else
'                SEAUTO = 1
'                SEPATHNO = (Mid(DATE11, 4, 2)) & "1"
'                UPDDAUTOPATH_QR(0) = SEPATHNO
'                UPDDAUTOPATH_QR(1) = Trim(dept)
'                UPDDAUTOPATH_QR(2) = Trim(SELDEPT_RS("MODALITYCD"))
'                UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
'                UPDDAUTOPATH_QR.Execute
'            End If
'        Else
'            SEAUTO = 1
'            SEPATHNO = Mid(DATE11, 4, 2) & "1"
'            INSDAUTOPATH_QR(0) = Trim(dept)
'            INSDAUTOPATH_QR(1) = Trim(SELDEPT_RS("MODALITYCD"))
'            'INSDAUTOPATH_QR(2) = Trim(Left(msk_caseno.FormattedText, 2))
'            INSDAUTOPATH_QR(2) = Trim(Left(MSFlexGrid1.TextMatrix(j, 2), 2))
'            INSDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
'            INSDAUTOPATH_QR(4) = SEPATHNO
'            INSDAUTOPATH_QR.Execute
'        End If
'Commented and added by Shraddha on 01/01/10
'********************************************
    Dim l As Integer
    If SELSEPATHNO_RS.RowCount > 0 Then
        
            'Next line added by Shraddha Pandit on 05/01/09
            l = Len(SELSEPATHNO_RS("PATHNO"))
           
            'Dim ST5, ST6 As Variant
            'ST5 = Left(SELSEPATHNO_RS("PATHNO"), 1)
            'ST6 = Trim(Mid(DATE11, 5, 1))
            ''' ADDED BY SANDHYA FOR SEROLOGY MONTHLY INITIALIZE,DATED 3/10/2008
            
            'next 999 if bet jan to sept
            If Trim(Mid(DATE11, 4, 2)) = "01" Or Trim(Mid(DATE11, 4, 2)) = "02" Or Trim(Mid(DATE11, 4, 2)) = "03" Or Trim(Mid(DATE11, 4, 2)) = "04" Or Trim(Mid(DATE11, 4, 2)) = "05" Or Trim(Mid(DATE11, 4, 2)) = "06" Or Trim(Mid(DATE11, 4, 2)) = "07" Or Trim(Mid(DATE11, 4, 2)) = "08" Or Trim(Mid(DATE11, 4, 2)) = "09" Then '999
                'Next line added by Shraddha Pandit on 05/01/09
                If Left(SELSEPATHNO_RS("PATHNO"), 1) = Trim(Mid(DATE11, 5, 1)) Then
                
                                  
                    'Next statement added by Shraddha Pandit on 05/01/09
                    If Len(Trim((SELSEPATHNO_RS("PATHNO")))) > 1 Then
                    SEAUTO = Trim(Mid(SELSEPATHNO_RS("PATHNO"), 2, l - 1))
                    Else
                    SEAUTO = Trim(SELSEPATHNO_RS("PATHNO"))
                    End If
                    
                    
                    
                    
                    'SEAUTO = Trim(SELSEPATHNO_RS("PATHNO"))
                       
                    'SEAUTO = Trim(Right(SELSEPATHNO_RS("PATHNO"), 4))
                    'SEAUTO = Mid(SELSEPATHNO_RS("PATHNO"), 3, 4)
                    'SEAUTO = Mid(SELSEPATHNO_RS("PATHNO"), 3)
                    SEAUTO = SEAUTO + 1
                         
                    'While Len(CStr(SEAUTO)) < 4  ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
                    '   SEAUTO = "0" & SEAUTO
                    'Wend
                    ''' SANDHYA, DATED 3/10/2008 , COMMENTED
                    'SEPATHNO = Mid(DATE11, 5, 1) & SEAUTO
                    'UPDDAUTOPATH_QR(0) = SEPATHNO
                    'Format of Path No is Month+pathno ex. 11 for jan+pathno as 1
                    'COMMENTED BY SHWETA 03-03-2014
                    UPDDAUTOPATH_QR(0) = (Mid(DATE11, 4, 2)) & SEAUTO
                    'UPDDAUTOPATH_QR(0) = SEAUTO
                    UPDDAUTOPATH_QR(1) = Trim(dept)
                    UPDDAUTOPATH_QR(2) = Trim(SELDEPT_RS("MODALITYCD"))
                    UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
                    UPDDAUTOPATH_QR.Execute
               Else
               
                   SEAUTO = 1
                   SEPATHNO = (Mid(DATE11, 4, 2)) & "1"
                   UPDDAUTOPATH_QR(0) = SEPATHNO
                   UPDDAUTOPATH_QR(1) = Trim(dept)
                        'SHWETA 01-06-2010
                   UPDDAUTOPATH_QR(2) = Trim(SELDEPT_RS("MODALITYCD")) 'Trim(MODALITY)
                   UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
                   UPDDAUTOPATH_QR.Execute
                   ''' USED FOR RESULTSET SANDHYA 3/10/2008
               End If
               ''' COMMENTED BY SANDHYA 3/10/2008
                        
            '*********** 999 if bet oct - Dec **********************
            ElseIf Trim(Mid(DATE11, 4, 2)) = "10" Or Trim(Mid(DATE11, 4, 2)) = "11" Or Trim(Mid(DATE11, 4, 2)) = "12" Then '999
               If Left(SELSEPATHNO_RS("PATHNO"), 2) = Trim(Mid(DATE11, 4, 2)) Then
                  
                   'SEAUTO = Trim(Mid(SELSEPATHNO_RS("PATHNO"), 2, l - 1))
                   SEAUTO = Trim(Mid(SELSEPATHNO_RS("PATHNO"), 3, l - 1)) 'Added on 02/11/09
                   SEAUTO = SEAUTO + 1
                   
                   UPDDAUTOPATH_QR(0) = (Mid(DATE11, 4, 2)) & SEAUTO
                   'SEAUTO = (Mid(DATE11, 4, 2)) & SEAUTO 'SHWETA 01-11-2012
                   UPDDAUTOPATH_QR(1) = Trim(dept)
                   UPDDAUTOPATH_QR(2) = Trim(SELDEPT_RS("MODALITYCD"))
                   UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
                   
                   UPDDAUTOPATH_QR.Execute
               Else
            
                   SEAUTO = 1
                   SEPATHNO = (Mid(DATE11, 4, 2)) & "1"
                   ' SEAUTO = SEPATHNO  'SHWETA 01-11-2012
                   UPDDAUTOPATH_QR(0) = SEPATHNO
                   UPDDAUTOPATH_QR(1) = Trim(dept)
                   UPDDAUTOPATH_QR(2) = Trim(SELDEPT_RS("MODALITYCD")) 'Trim(MODALITY)'changed by shweta 01-oct-2012
                   UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
                   UPDDAUTOPATH_QR.Execute
               End If
            End If
               ''' COMMENTED BY SANDHYA 3/10/2008
            '*************** 999 ******************
            
        Else
            
            SEAUTO = 1
            SEPATHNO = Mid(DATE11, 4, 2) & "1"
            INSDAUTOPATH_QR(0) = Trim(dept)
            INSDAUTOPATH_QR(1) = Trim(SELDEPT_RS("MODALITYCD"))
            INSDAUTOPATH_QR(2) = Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 2)
            INSDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            INSDAUTOPATH_QR(4) = SEPATHNO
            INSDAUTOPATH_QR.Execute
        End If
'********************************************
  '***** Comm and add "SR/" & Trim(Right(DATE11, 2)) & "/" to give new path no. for GZZ SE by Rahul at Sangrur On 15.12.2016
        'MSFlexGrid1.TextMatrix(j, 7) = SEAUTO
        MSFlexGrid1.TextMatrix(j, 7) = "SR/" & Trim(Right(DATE11, 2)) & "/" & (Mid(DATE11, 4, 2)) & "/" & SEAUTO
    End If
    End If
Next j

'***** Shraddha 02/12/2011 to show Sample description *****
            'MSFlexGrid1.Cols = 9
            MSFlexGrid1.Cols = 11
            MSFlexGrid1.ColWidth(8) = 5000
            ROW11 = MSFlexGrid1.row
            MSFlexGrid1.TextMatrix(0, 8) = "Test Description"
            
            str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
            Set SEL_DMUL_QR = rdb.CreateQuery("", str)

            SEL_DMUL_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
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
                'TESTDESC = TESTDESC & "| " & vbCrLf & Trim(SEL_MA101_RS("SER_DET"))
                'following line is added by Shweta on 07-05-2012
                ToEnterUnit Trim(SEL_DMUL_RS("PRO_CODE")), ROW11
                ''''till here
                SEL_DMUL_RS.MoveNext
            Wend
    
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 8) = TESTDESC
            
            '***********************************************************


Exit Sub
MSG:
MsgBox Err.Description
End Sub
'***** This ToEnterUnit function was Added by Shweta on 07-05-2012
Sub ToEnterUnit(ByVal procode As String, ByVal row1 As Integer)
MSFlexGrid1.row = row1 '- 1 'MSFlexGrid1.row - 1
MSFlexGrid1.col = 9 '- 1
TxtUnit.Width = MSFlexGrid1.CellWidth
TxtUnit.Height = MSFlexGrid1.CellHeight
TxtUnit.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
TxtUnit.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
'txt_pathno.Move MSFlexGrid1.CellLeft + MSFlexGrid1.Left + 4000, MSFlexGrid1.CellTop + MSFlexGrid1.Top
If procode = "F917" Or procode = "F053" Or procode = "F054" Or procode = "F977" Or procode = "F063" Or procode = "F915" Or procode = "F916" Or procode = "F918" Or procode = "F919" Or procode = "F920" Or procode = "F921" Or procode = "F922" Or procode = "F923" Or procode = "F924" Or procode = "F897" Or procode = "F898" Then
TxtUnit.Visible = True
Else
TxtUnit.Visible = False
End If
'
End Sub
Public Sub SPAUTONO_FED()

Dim UPD_ELCATEGORY_qr As rdoQuery
str = "UPDATE DISLIB.ELCATEGORY SET PATHNO=? WHERE TYPE='X' AND STATUS=?"
Set UPD_ELCATEGORY_qr = rdb.CreateQuery("", str)

Dim QR_ELCATEGORY As rdoQuery
Dim RS_ELCATEGORY As rdoResultset
str = "SELECT * FROM DISLIB.ELCATEGORY WHERE STATUS=? AND TYPE='X'"
Set QR_ELCATEGORY = rdb.CreateQuery("", str)


Dim PATH As String
Dim QR_HLTMTESTCODE As rdoQuery
Dim RS_HLTMTESTCODE As rdoResultset
str = "SELECT * FROM DISLIB.hltestcode WHERE ser_code=?"
Set QR_HLTMTESTCODE = rdb.CreateQuery("", str)
'On Error GoTo MSG:
I = 1
seldmulproce_qr1(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
Set seldmulproce_rs1 = seldmulproce_qr1.OpenResultset(1, 3)

If seldmulproce_rs1.EOF = False Then
   QR_HLTMTESTCODE(0) = Trim(seldmulproce_rs1("pro_code"))
Set RS_HLTMTESTCODE = QR_HLTMTESTCODE.OpenResultset(1, 3)
    If RS_HLTMTESTCODE.EOF = False Then
    TPY = RS_HLTMTESTCODE("type")
    End If
End If
I = MSFlexGrid1.row
'For i = 1 To MSFlexGrid1.Rows - 1
    If Trim(MSFlexGrid1.TextMatrix(I, 7)) = "" Then
        
        'QR_ELCATEGORY
        QR_ELCATEGORY(0) = TPY
        Set RS_ELCATEGORY = QR_ELCATEGORY.OpenResultset(1, 3)
    If RS_ELCATEGORY.EOF = False Then
        If Trim(Left(RS_ELCATEGORY("PATHNO"), 6)) = Trim(Format(DATE11, "YYMMDD")) Then
        Else
        If TPY = "CG" Then
            AUTOPATHNO = 801
        Else
            AUTOPATHNO = 401
        End If
        UPD_ELCATEGORY_qr(0) = Format(DATE11, "YYMMDD") & AUTOPATHNO
        UPD_ELCATEGORY_qr(1) = TPY
        
        UPD_ELCATEGORY_qr.Execute
        
            UPDDAUTOPATH_QR(0) = AUTOPATHNO
            UPDDAUTOPATH_QR(1) = "FED"
            UPDDAUTOPATH_QR(2) = TPY ' Trim(SELDEPT_RS("MODALITYCD"))
            UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            UPDDAUTOPATH_QR.Execute
        End If
    End If
            SELDAUTOPATH_QR(0) = "FED"
            'SELDAUTOPATH_QR(1) = Trim(MODALITY)
            SELDAUTOPATH_QR(1) = TPY ' Trim(SELDEPT_RS("MODALITYCD"))
            SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
            
            Set SELDAUTOPATH_RS = SELDAUTOPATH_QR.OpenResultset(2, 3)
        
        
        
        If SELDAUTOPATH_RS.RowCount > 0 Then
            AUTOPATHNO = SELDAUTOPATH_RS("PATHNO") + 1
            UPDDAUTOPATH_QR(0) = AUTOPATHNO
            UPDDAUTOPATH_QR(1) = "FED"
            UPDDAUTOPATH_QR(2) = TPY ' Trim(SELDEPT_RS("MODALITYCD"))
            UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            UPDDAUTOPATH_QR.Execute
        Else
            AUTOPATHNO = 1
            INSDAUTOPATH_QR(0) = "FED"
            INSDAUTOPATH_QR(1) = TPY 'Trim(MODALITYB)
            INSDAUTOPATH_QR(2) = Trim(Left(MSFlexGrid1.TextMatrix(I, 2), 2))
            INSDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            INSDAUTOPATH_QR(4) = 1
            INSDAUTOPATH_QR.Execute
        End If
        If Trim(dept) = "UZZ" Or Trim(dept) = "FED" Then
        SELDAUTOPATH_QR(0) = "FED"
            'SELDAUTOPATH_QR(1) = Trim(MODALITY)
            SELDAUTOPATH_QR(1) = TPY ' Trim(SELDEPT_RS("MODALITYCD"))
            SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
            
            Set SELDAUTOPATH_RS = SELDAUTOPATH_QR.OpenResultset(2, 3)
                PATH = Trim(SELDAUTOPATH_RS("PATHYEAR"))
        Else
            
            DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
            Set DPATHYEAR_RS = DPATHYEAR_QR.OpenResultset(1, 3)
            
            If DPATHYEAR_RS.RowCount > 0 Then
                PATH = DPATHYEAR_RS("LASTTWO")
            End If
        End If
    '***********padding the reqslno with leading zeros"****************
        Dim reqslno1
        reqslno1 = AUTOPATHNO
        If Trim(dept) = "FED" Then
            While Len(CStr(reqslno1)) < 4  ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
            ' NEEDED NUMBERING LIKE
            ' F05999999 FOR BF
            ' P05999999 FOR HL
            ' M05999999 FOR BM
            
                reqslno1 = "0" & reqslno1
            Wend
        Else
            While Len(CStr(reqslno1)) < 4
                reqslno1 = "0" & reqslno1
            Wend
        End If
        DAUTOPATHNO = reqslno1
'******************************************************************
        'If Trim(MODALITY) = "SP" Then
        If Trim(dept) = "UZZ" Or Trim(dept) = "FED" Then
            ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
            ' NEEDED NUMBERING LIKE
            ' F05999999 FOR BF
            ' P05999999 FOR HL
            ' M05999999 FOR BM
            'MSFlexGrid1.TextMatrix(I, 7) = Trim(PATH) & CStr(Trim(Right(DATE11, 2))) & Trim(DAUTOPATHNO)
            'Commented and added by Shraddha on 09/12/09
            'MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = Trim(PATH) & Replace(Format(DATE11, ddmmyyyy), "/", "") & Trim(DAUTOPATHNO)
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = Replace(Format(DATE11, "yymmdd"), "/", "") & Trim(DAUTOPATHNO)
        Else
            'MSFlexGrid1.TextMatrix(I, 7) = DAUTOPATHNO & "/" & PATH
            'Commented and added by Shraddha on 09/12/09
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = DAUTOPATHNO & "/" & PATH
        End If
        
        
        upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))
        upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
        upddrequisitn.Execute
        
        If Trim(Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1), 4, 2)) = "HL" Then
            upd_TEMMPCOLL_QR(0) = UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))
            upd_TEMMPCOLL_QR(1) = ipadd
            upd_TEMMPCOLL_QR(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
            upd_TEMMPCOLL_QR.Execute
            

        End If
        'End If
    End If
'Next i
Exit Sub
MSG:
MsgBox Err.Description
End Sub

'***** This SPAUTONO function was Added by Shraddha on 12/10/09
Public Sub SPAUTONO()
Dim PATH As String
'On Error GoTo MSG:
I = 1
For I = 1 To MSFlexGrid1.Rows - 1
    If Trim(MSFlexGrid1.TextMatrix(I, 5)) <> "" Then
        
            SELDAUTOPATH_QR(0) = Trim(dept)
            'SELDAUTOPATH_QR(1) = Trim(MODALITY)
            SELDAUTOPATH_QR(1) = Trim(SELDEPT_RS("MODALITYCD"))
            SELDAUTOPATH_QR(2) = Trim(Right(DATE11, 2))
            Set SELDAUTOPATH_RS = SELDAUTOPATH_QR.OpenResultset(2, 3)
        
        
        
        If SELDAUTOPATH_RS.RowCount > 0 Then
            AUTOPATHNO = SELDAUTOPATH_RS("PATHNO") + 1
            UPDDAUTOPATH_QR(0) = AUTOPATHNO
            UPDDAUTOPATH_QR(1) = Trim(dept)
            UPDDAUTOPATH_QR(2) = Trim(SELDEPT_RS("MODALITYCD"))
            UPDDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            UPDDAUTOPATH_QR.Execute
        Else
            AUTOPATHNO = 1
            INSDAUTOPATH_QR(0) = Trim(dept)
            INSDAUTOPATH_QR(1) = Trim(MODALITYB)
            INSDAUTOPATH_QR(2) = Trim(Left(MSFlexGrid1.TextMatrix(I, 2), 2))
            INSDAUTOPATH_QR(3) = Trim(Right(DATE11, 2))
            INSDAUTOPATH_QR(4) = 1
            INSDAUTOPATH_QR.Execute
        End If
        If Trim(dept) = "UZZ" Or Trim(dept) = "FED" Then
                PATH = Trim(SELDAUTOPATH_RS("PATHYEAR"))
        Else
            
            DPATHYEAR_QR(0) = "A" & Right(DATE11, 4)
            Set DPATHYEAR_RS = DPATHYEAR_QR.OpenResultset(1, 3)
            
            If DPATHYEAR_RS.RowCount > 0 Then
                PATH = DPATHYEAR_RS("LASTTWO")
            End If
        End If
    '***********padding the reqslno with leading zeros"****************
        Dim reqslno1
        reqslno1 = AUTOPATHNO
        If Trim(dept) = "UZZ" Or Trim(dept) = "FED" Then
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
        If Trim(dept) = "UZZ" Or Trim(dept) = "FED" Then
            ' CHANGED TO 6 FROM 7 AS HEMOTOLOGY
            ' NEEDED NUMBERING LIKE
            ' F05999999 FOR BF
            ' P05999999 FOR HL
            ' M05999999 FOR BM
            'MSFlexGrid1.TextMatrix(I, 7) = Trim(PATH) & CStr(Trim(Right(DATE11, 2))) & Trim(DAUTOPATHNO)
            'Commented and added by Shraddha on 09/12/09
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = Trim(PATH) & CStr(Trim(Right(DATE11, 2))) & Trim(DAUTOPATHNO)
        Else
            'MSFlexGrid1.TextMatrix(I, 7) = DAUTOPATHNO & "/" & PATH
            'Commented and added by Shraddha on 09/12/09
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = DAUTOPATHNO & "/" & PATH
        End If
        
        
        upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))
        upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
        upddrequisitn.Execute
        
        If Trim(Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1), 4, 2)) = "HL" Then
            upd_TEMMPCOLL_QR(0) = UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))
            upd_TEMMPCOLL_QR(1) = ipadd
            upd_TEMMPCOLL_QR(2) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
            upd_TEMMPCOLL_QR.Execute
            

        End If
        'End If
    End If
Next I
Exit Sub
MSG:
MsgBox Err.Description
End Sub
Public Sub FECPATHNO()



If MODALITYB = "BC" Or MODALITYB = "EL" Or MODALITYB = "TM" Then

'Increment and update for next records
    'PathNo=pathno+1 7th col
    ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
    pathno = MSFlexGrid1.TextMatrix(MSFlexGrid1.row - 1, 7) 'Clear it on NEW Load
    If MSFlexGrid1.row > 1 Then
        'pathno = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)
        
        pathno = pathno + 1
        
        '***** Shraddha 05/11/09 *****
        Dim zeros As String
        zeros = ""
        'While Len(CStr(pathno)) < 3
         
'            If Len(CStr(pathno)) = 1 Then
'                zeros = "00" & CStr(pathno)
'            ElseIf Len(CStr(pathno)) = 2 Then
'                zeros = "0" & CStr(pathno)
'            Else
'                zeros = pathno
'            End If
            
            '***** Commented and added by Shraddha 03/02/10 *****
            If Len(CStr(pathno)) = 1 Then
                zeros = "000" & CStr(pathno)
            ElseIf Len(CStr(pathno)) = 2 Then
                zeros = "00" & CStr(pathno)
            ElseIf Len(CStr(pathno)) = 3 Then
                zeros = "0" & CStr(pathno)
            Else
                zeros = pathno
            End If
            '****************************************************
        'Wend
        '*****************************
        
    End If
   
        
        '*****************************
        '***** Shraddha 10/02/10 *****
        '***** If PathNo already given 16/10/09 *****
    '*****
    Dim MODALITY As String
    Dim hlreqno As String
    SELDEPT_QR(0) = MODALITYB
    Set SELDEPT_RS = SELDEPT_QR.OpenResultset(1, 3)
    MODALITY = Trim(SELDEPT_RS("MODALITYCD"))
    
    'If Trim(txt_pathno.Text) <> "" Then
    If zeros <> "" Then
        If MODALITY = "BC" Or MODALITY = "EL" Or MODALITY = "TM" Then
            For I = 1 To MSFlexGrid1.Rows - 1
                If Trim(txt_pathno.Text) = Trim(MSFlexGrid1.TextMatrix(I, 6)) Then
                    MsgBox "Lab No. Already exists"
                    Exit Sub
                End If
            Next
        
            'If Len(Trim(txt_pathno.Text)) <= 4 And Len(Trim(txt_pathno.Text)) <> 0 Then
            If Len(Trim(zeros)) <= 4 And Len(Trim(zeros)) <> 0 Then
                SELDREQUISITNQR(0) = ST & UCase(Trim(zeros))
            End If
            If Len(Trim(txt_pathno.Text)) > 4 Then
                SELDREQUISITNQR(0) = UCase(Trim(zeros))
            End If
            SELDREQUISITNQR(1) = Trim(dept) & MODALITY & "%"
            Set SELDREQUISITNRS = SELDREQUISITNQR.OpenResultset(1, 3)
        End If

            If SELDREQUISITNRS.RowCount > 0 Then
                'MsgBox "The Path No. is  Given already " & SELDREQUISITNRS("caseno") & vbCrLf & "Sample not yet received, Please give different Path No"
                MsgBox "The Path No. is  Given already " & SELDREQUISITNRS("caseno") & " " & SELDREQUISITNRS("REQNO") & vbCrLf & "Sample not yet received, Please give different Path No"
                'txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                txt_pathno.Visible = True '11/12/09
                txt_pathno.Top = 3480
                txt_pathno.Left = 12720
                txt_pathno.Text = ""
                
                MSFlexGrid1.clear

                MSFlexGrid1.Rows = 2
                           
                
                'TXTREQNO_Change
                Exit Sub
            End If
    End If

'*****
        '*****************************
        
    'MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = pathno
        
    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = zeros
    
'    If Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))) > 3 Then
'        upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))
'    End If
    'MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = txt_pathno.Text
    If Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))) <= 4 And Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))) <> 0 Then
        upddrequisitn(0) = ST & UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))
        upd_TEMMPCOLL_QR(0) = ST & UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))
    End If
    upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
    upd_TEMMPCOLL_QR(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
    upddrequisitn.Execute
    upd_TEMMPCOLL_QR.Execute
        
    txt_pathno.Visible = False

    
End If
End Sub

Private Sub cmb_location_LostFocus()
If cmb_location.Text = "" Then
    MsgBox "Please Select Location .......!"
    'cmb_location.SetFocus
Else
    If Trim(cmb_location.Text) = "HAEMATOPATHOLOGY LABORATORY$UZZ" Then
        cmd_smplrcvPRINT.Enabled = True
         
        Dim TREC_QR As rdoQuery
Dim TREC_RS As rdoResultset
str = "select * from DISLIB.TEMPRECEIV WHERE modalitycd='HL' and  sentflag='' and RECDIPAD='" & Trim(Winsock1.LocalIP) & "'"
Set TREC_QR = rdb.CreateQuery("", str)
Set TREC_RS = TREC_QR.OpenResultset(2, 3)
If TREC_RS.EOF = False Then
MsgBox "Please take Print out of Pending list"
cmd_pendlist_Click
End If

        
    Else
        cmd_smplrcvPRINT.Enabled = False
    End If

    TXTREQNO.SetFocus
End If
End Sub

Private Sub cmd_back_Click()

If Trim(cmb_location.Text) = "HAEMATOPATHOLOGY LABORATORY$UZZ" Then
If flg = False Then
    MsgBox "Please take the print of List of received samples"
    Exit Sub
End If
End If


flg = True
'Me.Hide

If showmain = True Then
   Me.MousePointer = 11
 '   Me.Hide
    collectionNEW.show
    Me.MousePointer = 0
       
Else
    Me.MousePointer = 11
 '   Me.Hide
    DIS_MAIN.show
    Me.MousePointer = 0
End If

'Dim X As rdoQuery
'For Each X In Barcode.VA
'Set X = Nothing
'Next

showmain = False
Unload Me
End Sub
'Shraddha Pandit 1/12/08
Public Sub deptdesc()

'to select color for printing on the barcode label

colorcd = ""
    If smplcd = "A" Or smplcd = "B" Then
        colorcd = "RED"
    ElseIf smplcd = "R" Or smplcd = "S" Or smplcd = "P" Then
        colorcd = "GRAY"
    ElseIf smplcd = "U" Then
        colorcd = "BLUE"
    ElseIf smplcd = "C" Then
        colorcd = "PURPLE"
    ElseIf smplcd = "D" Then
        colorcd = "BLACK"
    Else
        colorcd = ""
    End If
End Sub


Private Sub cmd_pendlist_Click()
pend = True
'cmd_smplrcvPRINT_Click
For I = 0 To 10
    CrystalReport1.Formulas(I) = ""
Next
    
stamp = Right(UID, 6) & " - " & ipadd & " - " & Format(DATE11, "DD/MM/YYYY") & " - " & time1

'CrystalReport1.SelectionFormula = "{TEMPCOLLCT.MODALITYCD} = 'HL' AND {TEMPCOLLCT.FLD1} <> ''  AND {TEMPCOLLCT.RECDIPAD}='" & ipadd & "'"
'Commented and added by Shraddha on 18/10/2011
'ipadd = "10.1.4.90"
CrystalReport1.SelectionFormula = "{TEMPRECEIV.MODALITYCD} = 'HL' AND  {TEMPRECEIV.SENTFLAG} = '' AND {TEMPRECEIV.RECDIPAD}='" & ipadd & "'" ' AND {TEMPRECEIV.SENTDATE} in Date (" & Format(Trim(Date), "YYYY, MM, DD") & ")"

'CrystalReport1.SelectionFormula = "{TEMPCOLLCT.MODALITYCD} = 'HL' AND  {TEMPCOLLCT.SENTFLAG} = ''" 'AND {TEMPCOLLCT.RECDIPAD}='" & ipadd & "'"
CrystalReport1.Formulas(0) = "STAMP = '" & stamp & "'"
'CrystalReport1.Formulas(1) = "MODALITY = '" & MODALITY & "'"
CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\SampleRcvHL.RPT"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Destination = crptToPrinter

CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
 CrystalReport1.Action = 1
str = "UPDATE DISLIB.TEMPRECEIV SET SENTFLAG='P',RECDFLAG='R' WHERE RECDIPAD='" & Trim(Winsock1.LocalIP) & "'"
rdb.Execute str
'TXTREQNO.SetFocus
End Sub

Private Sub cmd_print_Click()
'*********** Added on 1st Dec'08 by Shraddha Pandit *************

Dim QTY1 As String
Dim BAR_CD As String
Dim EXP_DT As String
Dim ITEM_DESC As String
Dim TEMPDATA As String
Dim rack As String
Dim batchno1 As String
Dim itemcode1 As String
Dim I As Long, START
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

'------------------------------
Dim SELLABNO_QR As rdoQuery
Dim SELLABNO_RS As rdoResultset


'------------------------------
'***********************************
Dim SELDEPTDES_QR As rdoQuery
Dim SELDEPTDES_RS As rdoResultset
Dim VIEWFLAG_BCD As Boolean
Dim NEWFLAG_BCD As Boolean
Dim colorcd As String
Dim smplcd As String
Dim DESCRIP As String
Dim UPDRTIMEFORBC As rdoQuery

flg = False
'************************************

'For Barcode 1/12/08

    str = "SELECT * FROM PATOOLSLIB.MA106 WHERE DEPT_CODE=?"
    Set SELDEPTDES_QR = rdb.CreateQuery("", str)

'If Mid(reqno, 4, 2) = "SE" Then
'for staff
If Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))) = 6 Then
    f1 = FreeFile()
    Open "c:\micro.TXT" For Output As #f1
    Dim rs_emp As rdoResultset
    Dim DPT1 As String

    str = "SELECT distinct SAMPLECODE.SAMPLECD FROM DISLIB.SAMPLECODE,DISLIB.DMULPROCE WHERE DMULPROCE.PRO_CODE=SAMPLECODE.SERCODE AND DMULPROCE.REQNO=?"
    Set SELLABNO_QR = rdb.CreateQuery("", str)

    SELLABNO_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
    SELLABNO_QR.Execute
    

    Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
    Dim cstaff As Integer
    cstaff = SELLABNO_RS.RowCount

    While Not SELLABNO_RS.EOF
        'Dim REQ111t, P_CODE111t, SCODE111t
        'REQ111t = reqno
        REQ111t = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
                        
        SCODE111t = SELLABNO_RS(0)
        smplcd = SELLABNO_RS(0)
                        
        REQ22 = Mid$(REQ111t, 4)
        SCODE111t = REQ22 & SCODE111t
                        
        deptdesc
                        
        DPT1 = Trim(Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1), 1, 3))
        
        SELDEPTDES_QR(0) = DPT1
        Set SELDEPTDES_RS = SELDEPTDES_QR.OpenResultset(1, 3)
        
        DESCRIP = ""
        DESCRIP = Mid(Trim(SELDEPTDES_RS("DEPT_NAME")), 1, 8) & "/" & colorcd
                           
        Set rs_emp = rdb.OpenResultset("select age,sex from dislib.dreqstaffd where empcode='" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)) & "'", 1, 1)
        agesex1 = rs_emp(0) & "/" & rs_emp(1)
                       
        PRNTSTR = "I8,1,001" & vbCrLf & "q591" & vbCrLf & "S3" & vbCrLf & "ON,D" & vbCrLf & "JF" & vbCrLf & "j189" & vbCrLf & "WN" & vbCrLf & "D10" & vbCrLf & "ZB" & vbCrLf & "Q217,44" & vbCrLf & "N" & vbCrLf
        'PRNTSTR1 = PRNTSTR & "A52,11,0,2,1,1,N,""" & Trim(txt_name) & """" & vbCrLf
        PRNTSTR1 = PRNTSTR & "A52,11,0,2,1,1,N,""" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)) & """" & vbCrLf
        PRNTSTR1 = PRNTSTR1 & "B92,41,0,1,3,6,79,N,""" & SCODE111t & """" & vbCrLf
        PRNTSTR1 = PRNTSTR1 & "A194,124,0,2,1,1,N,""" & SCODE111t & """" & vbCrLf
        PRNTSTR1 = PRNTSTR1 & "A52,152,0,3,1,1,N,""" & agesex1 & """" & vbCrLf
        'PRNTSTR1 = PRNTSTR1 & "A150,150,0,4,1,1,N,""" & Trim(msk_caseno) & """" & vbCrLf
        PRNTSTR1 = PRNTSTR1 & "A150,150,0,4,1,1,N,""" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)) & """" & vbCrLf
        PRNTSTR1 = PRNTSTR1 & "A340,149,0,2,1,1,N,""" & DESCRIP & """" & vbCrLf
                
        PRNTSTR1 = PRNTSTR1 & "P1" & vbCrLf
        PRNTSTR1 = PRNTSTR1 & "N" & vbCrLf
    
        Print #f1, PRNTSTR1
    
        SELLABNO_RS.MoveNext
    Wend
    Close #f1
    Shell "c:\DISREPORTSEC\DATA1\micro.bat", 0
    'SELPROCODE_QR.Close
    SELLABNO_QR.Close
    'Exit Sub
Else
    'Kill "C:\MICRO.TXT"
    f1 = FreeFile()
    Open "c:\micro.TXT" For Output As #f1
    Dim rs_ta101 As rdoResultset

    str = "SELECT distinct SAMPLECODE.SAMPLECD FROM DISLIB.SAMPLECODE,DISLIB.DMULPROCE WHERE DMULPROCE.PRO_CODE=SAMPLECODE.SERCODE AND DMULPROCE.REQNO=?"
    Set SELLABNO_QR = rdb.CreateQuery("", str)

    SELLABNO_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
    SELLABNO_QR.Execute

    Set SELLABNO_RS = SELLABNO_QR.OpenResultset(1, 3)
    Dim c As Integer
    c = SELLABNO_RS.RowCount

    While Not SELLABNO_RS.EOF
        'Dim REQ111t, P_CODE111t, SCODE111t
        REQ111t = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))

        SCODE111t = SELLABNO_RS(0)
        smplcd = SELLABNO_RS(0)
                        
        REQ22 = Mid$(REQ111t, 4)
        SCODE111t = REQ22 & SCODE111t
                                    
        deptdesc
        
        DPT = Trim(Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1), 1, 3))
                    
        SELDEPTDES_QR(0) = DPT
        Set SELDEPTDES_RS = SELDEPTDES_QR.OpenResultset(1, 3)
        
        
        DESCRIP = ""
        DESCRIP = Mid(Trim(SELDEPTDES_RS("DEPT_NAME")), 1, 8) & "/" & colorcd
                        
        'Set rs_ta101 = rdb.OpenResultset("select birth_date,sex from patoolslib.ta101 where case_no='" & Trim(msk_caseno.FormattedText) & "'", 1, 1)
        Set rs_ta101 = rdb.OpenResultset("select birth_date,sex from patoolslib.ta101 where case_no='" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)) & "'", 1, 1)
        Dim n As Integer
        n = DateDiff("m", rs_ta101("BIRTH_DATE"), CStr(DATE11)) / 12
        agesex = n & "/" & rs_ta101(1)
                            
        PRNTSTR = "I8,1,001" & vbCrLf & "q591" & vbCrLf & "S3" & vbCrLf & "ON,D" & vbCrLf & "JF" & vbCrLf & "j189" & vbCrLf & "WN" & vbCrLf & "D10" & vbCrLf & "ZB" & vbCrLf & "Q217,44" & vbCrLf & "N" & vbCrLf
                        
        'PRNTSTR1 = PRNTSTR & "A52,11,0,2,1,1,N,""" & Trim(txt_name) & """" & vbCrLf
        PRNTSTR1 = PRNTSTR & "A52,11,0,2,1,1,N,""" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)) & """" & vbCrLf
        PRNTSTR1 = PRNTSTR1 & "B92,41,0,1,3,6,79,N,""" & SCODE111t & """" & vbCrLf
        PRNTSTR1 = PRNTSTR1 & "A194,124,0,2,1,1,N,""" & SCODE111t & """" & vbCrLf
        PRNTSTR1 = PRNTSTR1 & "A52,152,0,3,1,1,N,""" & agesex & """" & vbCrLf
        'PRNTSTR1 = PRNTSTR1 & "A150,150,0,4,1,1,N,""" & Trim(msk_caseno) & """" & vbCrLf
        PRNTSTR1 = PRNTSTR1 & "A150,150,0,4,1,1,N,""" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)) & """" & vbCrLf
        PRNTSTR1 = PRNTSTR1 & "A340,149,0,2,1,1,N,""" & DESCRIP & """" & vbCrLf
        PRNTSTR1 = PRNTSTR1 & "P1" & vbCrLf
        PRNTSTR1 = PRNTSTR1 & "N" & vbCrLf
        Print #f1, PRNTSTR1
                        
        SELLABNO_RS.MoveNext
                  
    Wend
    Close #f1
    Shell "c:\DISREPORTSEC\DATA1\micro.bat", 0
    'SELPROCODE_QR.Close
    SELLABNO_QR.Close
    'Exit Sub
End If
End Sub

Private Sub cmd_print_report_Click()
'Dim Modlt, RecDate
'MSFlexGrid1.row = 1
'
'Modlt = MSFlexGrid1.TextMatrix(1, 4)
'RecDate = MSFlexGrid1.TextMatrix(1, 5)
'
'        var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
'        var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
'
'        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DEPTCOLLECTION.RPT"
'        CrystalReport1.SelectionFormula = "{RREQUISITN.EXAMDT} >= Date(" & var_date_f & ") AND {RREQUISITN.EXAMDT} <= Date(" & var_date_t & ") and {RREQUISITN.DEPT_CODE}= '" & dept & " '  AND {RREQUISITN.MODALITYCD}= '" & MODALITY & "' "
'        CrystalReport1.Formulas(0) = "FROMDT = '" & msk_fromdt.FormattedText & "'"
'        CrystalReport1.Formulas(1) = "TODT = '" & msk_todt.FormattedText & "'"
'        CrystalReport1.Formulas(2) = "MODALITY = '" & Left(cmb_modality.Text, 30) & "'"
'        CrystalReport1.Formulas(3) = "USRID = '" & str & "'"
       
        
'***** Shraddha 31/05/10 *****
'var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
'var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)

flg = False

var_date = Format(DATE11, "DD/MM/YYYY")

var_date = Year(var_date) & "," & month(var_date) & "," & Day(var_date)

CrystalReport1.ReportFileName = "C:\disreportsec\HaematDailyDeptRcv1.rpt"
CrystalReport1.Formulas(0) = "FROMDT='" & Trim(var_date) & "'"
CrystalReport1.Formulas(2) = "TODT='" & Trim(var_date) & "'"
'CrystalReport1.SelectionFormula = "{RREQUISITN.Freportdt}>=DATE(" & var_date & ")AND {RREQUISITN.Freportdt}<=DATE(" & var_date & ") AND ({RREQUISITN.REQNSTAT}='F' or {RREQUISITN.REQNSTAT}='T') AND {RREQUISITN.DEPT_CODE}='UZZ' AND {RREQUISITN.MODALITYCD}='HL'"
'CrystalReport1.SelectionFormula = "{RTIMELOG.SAMRECDATE}= DATE(" & var_date & ") AND {RREQUISITN.MODALITYCD}='HL'"
'CrystalReport1.SelectionFormula = "{RTIMELOG.SAMRECDATE}= DATE(" & var_date & ") "
'CrystalReport1.SelectionFormula = "{HEMDAILYLF.SAMRECDATE}= DATE(" & var_date & ") AND {HEMDAILYLF.MODALITYCD}='HL'"
CrystalReport1.SelectionFormula = "{HEMDAILYLF.SAMRECDATE}= DATE(" & var_date & ")"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Destination = crptToWindow
CrystalReport1.Action = 1
'*****************************
        
End Sub

'Private Sub cmd_print_report_Click()
'CR.ReportFileName = "C:\DISREPORTSEC\micro_icu1.rpt"
'    CR.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'    UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)
'    CR.Formulas(0) = "from=Date(" & var_date_f & ")"
'    CR.Formulas(1) = "to=Date(" & var_date_t & ")"
'    CR.Formulas(2) = "USRID = '" & UI & "'"

'    CR.SelectionFormula = "{PATHOGEN.REQNDT}>=Date(" & var_date_f & ") AND {PATHOGEN.REQNDT}<=Date(" & var_date_t & ") and {PATHOGEN.refbydoc}='" & doctor & "'"

'    CR.WindowState = crptMaximized
'    CR.Action = 1
'End Sub

Private Sub cmd_refresh_Click()

If Trim(cmb_location.Text) = "HAEMATOPATHOLOGY LABORATORY$UZZ" Then
If flg = False Then
'COMMENTED BY SHWETA AT SANGRUR 26-08-2015
    'MsgBox "Please take the print of List of received samples"
   ' Exit Sub
End If
End If
'SHWETA 13-FEB-2012 COMMENTED  AS TOLD BY KALPANA MAM
'***** Shraddha 09/11/2011 *****
'update to P
'upd_TCOLL_P_QR(0) = ipadd
'upd_TCOLL_P_QR.Execute
'*******************************

flg = True
MSFlexGrid1.clear
'cmb_location.Text = ""
MSFlexGrid1.Rows = 2
'cmb_location.Text = "PVT. OPD"

'cmb_location.Enabled = False 'Shraddha Pandit 27/11/08
TXTREQNO.Text = ""
TXTREQNO.SetFocus
txt_pathno.Text = ""
txt_pathno.Visible = False
'cmb_location.SetFocus
End Sub


Private Sub cmd_smplrcvPRINT_Click()

'***** Shraddha 09/08/2011 *****

If Trim(MSFlexGrid1.TextMatrix(1, 1)) = "" Then
    MsgBox "Please scan the samples."
    Exit Sub
End If

Dim u

Dim X
X = 1

While (X < MSFlexGrid1.Rows)

    SEL_TEMPCOLL_QR(0) = MSFlexGrid1.TextMatrix(X, 1)
    Set SEL_TEMPCOLL_RS = SEL_TEMPCOLL_QR.OpenResultset(1, 3)

    If Not SEL_TEMPCOLL_RS.EOF Then
        If Trim(SEL_TEMPCOLL_RS("SENTFLAG")) <> "P" Then
            upd_TCOLL_blank_QR(0) = MSFlexGrid1.TextMatrix(X, 1)
            upd_TCOLL_blank_QR.Execute
            
            upd_TCOLL_RCVIP_QR(0) = ipadd
            upd_TCOLL_RCVIP_QR(1) = MSFlexGrid1.TextMatrix(X, 1)
            upd_TCOLL_RCVIP_QR.Execute
        End If
    End If

    
'    MSFlexGrid1.row = X - 1
'    TEMPRECEIVE
    X = X + 1
Wend

'*******************************


''*******************************

'
''*******************************

    
For I = 0 To 10
    CrystalReport1.Formulas(I) = ""
Next
    
stamp = Right(UID, 6) & " - " & ipadd & " - " & Format(DATE11, "DD/MM/YYYY") & " - " & time1

'CrystalReport1.SelectionFormula = "{TEMPCOLLCT.MODALITYCD} = 'HL' AND {TEMPCOLLCT.FLD1} <> ''  AND {TEMPCOLLCT.RECDIPAD}='" & ipadd & "'"
'Commented and added by Shraddha on 18/10/2011
'ipadd = "10.1.4.90"
CrystalReport1.SelectionFormula = "{TEMPRECEIV.MODALITYCD} = 'HL' AND  {TEMPRECEIV.SENTFLAG} = '' AND {TEMPRECEIV.RECDIPAD}='" & ipadd & "'" ' AND {TEMPRECEIV.SENTDATE} in Date (" & Format(Trim(Date), "YYYY, MM, DD") & ")"

'CrystalReport1.SelectionFormula = "{TEMPCOLLCT.MODALITYCD} = 'HL' AND  {TEMPCOLLCT.SENTFLAG} = ''" 'AND {TEMPCOLLCT.RECDIPAD}='" & ipadd & "'"
CrystalReport1.Formulas(0) = "STAMP = '" & stamp & "'"
'CrystalReport1.Formulas(1) = "MODALITY = '" & MODALITY & "'"
CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\SampleRcvHL.RPT"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
 CrystalReport1.Action = 1

'**************************
''CrystalReport3.SelectionFormula = "{TEMPCOLLCT.MODALITYCD} = 'HL' AND {TEMPCOLLCT.FLD1} <> '' AND  {TEMPCOLLCT.SENTFLAG} = '' AND {TEMPCOLLCT.RECDIPAD}='" & ipadd & "'"
''Commented and added by Shraddha on 09/11/2011
'CrystalReport3.SelectionFormula = "{TEMPCOLLCT.MODALITYCD} = 'HL' AND  {TEMPCOLLCT.SENTFLAG} = '' AND {TEMPCOLLCT.RECDIPAD}='" & ipadd & "'"
'CrystalReport3.Formulas(0) = "STAMP = '" & stamp & "'"
'CrystalReport3.ReportFileName = "C:\DISREPORTSECLF\SampleRcvHL_GEN.RPT"
'CrystalReport3.WindowState = crptMaximized
'CrystalReport3.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'CrystalReport3.Action = 1
''**************************

'TEMPRECEIVE
''''''
str = "UPDATE DISLIB.TEMPRECEIV SET SENTFLAG='P',RECDFLAG='R' WHERE RECDIPAD='" & Trim(Winsock1.LocalIP) & "'"
rdb.Execute str
''''''''''''
u = 1

While (u < MSFlexGrid1.Rows)
       
    'dept = Trim(Left(MSFlexGrid1.TextMatrix(u, 1), 3))
    'MODCD = Trim(Mid(MSFlexGrid1.TextMatrix(u, 1), 4, 2))
       
    upd_TCOLL_Print_QR(0) = MSFlexGrid1.TextMatrix(u, 1)
    upd_TCOLL_Print_QR.Execute
    
    u = u + 1
Wend

flg = True 'Shraddha 16/08/2011

End Sub
Sub TEMPRECEIVE()
Dim LP As Integer

'On Error GoTo MSG11



'FOR SRNO INCREMENT
Dim srno As Integer
Dim TREC_QR As rdoQuery
Dim TREC_RS As rdoResultset
str = "select MAX(SRNO) from DISLIB.TEMPRECEIV WHERE PRLSTDT=CURRENT DATE"
Set TREC_QR = rdb.CreateQuery("", str)
Set TREC_RS = TREC_QR.OpenResultset(2, 3)
If IsNull(TREC_RS(0)) = True Then
srno = 0
Else
srno = TREC_RS(0)

End If


Set TREC_QR = Nothing
Set TREC_RS = Nothing

'SRNO,REQNO,MODALITYCD,CASENO,NAME,SEX,
'AGE,IPADD,PATNSTAT,PAT_CATG,LOCATION,COLLDATE,COLLTIME,
'SER_CODE,SER_DET,SENTFLAG,USERID,FLD1,RECDDATE,RECDTIME,
'RECDFLAG,RECDIPAD,PRLSTDT,PRLSTTM
'For LP = 1 To MSFlexGrid1.Rows - 2
'already exist check
LP = MSFlexGrid1.row '- 1

    str = "SELECT * FROM Dislib.TEMPRECEIV WHERE REQNO=?" 'DREQSTAFFD
    Set SEL_TA101_QR = rdb.CreateQuery("", str)
    Set TREC_QR = rdb.CreateQuery("", str)
    TREC_QR(0) = Trim(MSFlexGrid1.TextMatrix(LP, 1))
    Set TREC_RS = TREC_QR.OpenResultset(2, 3)
    If TREC_RS.EOF = False Then
    'GoTo abc
    Exit Sub
    End If

Dim INSTEMPREC_QR As rdoQuery
str = "INSERT INTO DISLIB.TEMPRECEIV(SRNO,REQNO,MODALITYCD,CASENO,NAME,SEX,AGE,IPADD,PATNSTAT,PAT_CATG,LOCATION,COLLDATE,COLLTIME,SER_CODE,SER_DET,SENTFLAG,USERID,FLD1,RECDDATE,RECDTIME,RECDFLAG,RECDIPAD,PRLSTDT,PRLSTTM,FLD3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
Set INSTEMPREC_QR = rdb.CreateQuery("", str)


'TA101
srno = srno + 1
If Len(Trim(MSFlexGrid1.TextMatrix(LP, 2))) = 6 Then
    str = "SELECT * FROM Dislib.DREQSTAFFD WHERE REQNO=?" 'DREQSTAFFD
    Set SEL_TA101_QR = rdb.CreateQuery("", str)
    SEL_TA101_QR(0) = Trim(MSFlexGrid1.TextMatrix(LP, 1))
Else
    str = "select * from PATOOLSLIB.ta101 where case_no=?"
    Set SEL_TA101_QR = rdb.CreateQuery("", str)
    SEL_TA101_QR(0) = Trim(MSFlexGrid1.TextMatrix(LP, 2))
End If

age = ""
SEX = ""
Set SEL_TA101_RS = SEL_TA101_QR.OpenResultset(2, 3)
If Len(Trim(MSFlexGrid1.TextMatrix(LP, 2))) = 6 Then
'age = DateDiff("YYYY", SEL_TA101_RS("AGE"), CStr(DATE11))
age = SEL_TA101_RS("AGE")
SEX = Trim(SEL_TA101_RS("SEX"))
pat_catg = "ST"

Else
age = DateDiff("YYYY", SEL_TA101_RS("BIRTH_DATE"), CStr(DATE11))
SEX = Trim(SEL_TA101_RS("SEX"))
pat_catg = Trim(SEL_TA101_RS("PAT_CATG"))

End If
Dim dmg1 As String
If Len(Trim(MSFlexGrid1.TextMatrix(LP, 2))) = 6 Then
dmg1 = ""
Else
dmg1 = Trim(SEL_TA101_RS("UD_NAME"))
End If



INSTEMPREC_QR(0) = srno
INSTEMPREC_QR(1) = MSFlexGrid1.TextMatrix(LP, 1) 'REQNO
INSTEMPREC_QR(2) = Mid(Trim(MSFlexGrid1.TextMatrix(LP, 1)), 4, 2) 'MODALITYCD
INSTEMPREC_QR(3) = MSFlexGrid1.TextMatrix(LP, 2) 'CASENO
INSTEMPREC_QR(4) = MSFlexGrid1.TextMatrix(LP, 3) 'NAME
INSTEMPREC_QR(5) = SEX 'SEX
INSTEMPREC_QR(6) = age 'AGE
INSTEMPREC_QR(7) = "" 'IPADD (COLLECTION)
INSTEMPREC_QR(8) = MSFlexGrid1.TextMatrix(LP, 10) 'PAT_STAT
INSTEMPREC_QR(9) = pat_catg 'PAT_CATG
INSTEMPREC_QR(10) = cmb_location.Text 'LOCATION
INSTEMPREC_QR(11) = GET_REVDATE(MSFlexGrid1.TextMatrix(LP, 5)) 'COLLDATE
INSTEMPREC_QR(12) = MSFlexGrid1.TextMatrix(LP, 6) 'COLLTIME
INSTEMPREC_QR(13) = "test"   'SER_CODE
INSTEMPREC_QR(14) = MSFlexGrid1.TextMatrix(LP, 9) 'SER_DET
INSTEMPREC_QR(15) = ""     'SENT FLAG
INSTEMPREC_QR(16) = Trim(logonfrm.usrid) 'USERID
INSTEMPREC_QR(17) = MSFlexGrid1.TextMatrix(LP, 7) 'FLD1 (PATHNO)
INSTEMPREC_QR(18) = GET_REVDATE(DATE11)  'RECDDATE
INSTEMPREC_QR(19) = time1  'RECTIME
INSTEMPREC_QR(20) = dmg1 'RECDFLAG
INSTEMPREC_QR(21) = Winsock1.LocalIP  'RECDIAPP
INSTEMPREC_QR(22) = GET_REVDATE(DATE11)  'PR DATE
INSTEMPREC_QR(23) = time1  'PR TIME
INSTEMPREC_QR(24) = MSFlexGrid1.TextMatrix(LP, 8) 'OPPROCE
INSTEMPREC_QR.Execute
abc:
'Exit Sub
''Next
'MSG11:
'MsgBox "ERROR : " & Err.Description & vbCrLf & "TAKE SCREENSHOT AND CONTACT TO COMPUTER DEPT."
'End
End Sub

Sub TEMPRECEIVE_old()






'FOR SRNO INCREMENT
Dim srno As Integer
Dim TREC_QR As rdoQuery
Dim TREC_RS As rdoResultset
str = "select MAX(SRNO) from DISLIB.TEMPRECEIV WHERE PRLSTDT=CURRENT DATE"
Set TREC_QR = rdb.CreateQuery("", str)
Set TREC_RS = TREC_QR.OpenResultset(2, 3)
If IsNull(TREC_RS(0)) = True Then
srno = 0
Else
srno = TREC_RS(0)

End If


Set TREC_QR = Nothing
Set TREC_RS = Nothing

Dim INSTEMPREC_QR As rdoQuery, LP As Integer
str = "INSERT INTO DISLIB.TEMPRECEIV(SRNO,REQNO,MODALITYCD,CASENO,NAME,SEX,AGE,IPADD,PATNSTAT,PAT_CATG,LOCATION,COLLDATE,COLLTIME,SER_CODE,SER_DET,SENTFLAG,USERID,FLD1,RECDDATE,RECDTIME,RECDFLAG,RECDIPAD,PRLSTDT,PRLSTTM,FLD3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
Set INSTEMPREC_QR = rdb.CreateQuery("", str)
'SRNO,REQNO,MODALITYCD,CASENO,NAME,SEX,
'AGE,IPADD,PATNSTAT,PAT_CATG,LOCATION,COLLDATE,COLLTIME,
'SER_CODE,SER_DET,SENTFLAG,USERID,FLD1,RECDDATE,RECDTIME,
'RECDFLAG,RECDIPAD,PRLSTDT,PRLSTTM
For LP = 1 To MSFlexGrid1.Rows - 2
'already exist check

    str = "SELECT * FROM Dislib.TEMPRECEIV WHERE REQNO=?" 'DREQSTAFFD
    Set SEL_TA101_QR = rdb.CreateQuery("", str)
    Set TREC_QR = rdb.CreateQuery("", str)
    TREC_QR(0) = Trim(MSFlexGrid1.TextMatrix(LP, 1))
    Set TREC_RS = TREC_QR.OpenResultset(2, 3)
    If TREC_RS.EOF = False Then
    GoTo abc
    End If



'TA101
srno = srno + 1
If Len(Trim(MSFlexGrid1.TextMatrix(LP, 2))) = 6 Then
    str = "SELECT * FROM Dislib.DREQSTAFFD WHERE REQNO=?" 'DREQSTAFFD
    Set SEL_TA101_QR = rdb.CreateQuery("", str)
    SEL_TA101_QR(0) = Trim(MSFlexGrid1.TextMatrix(LP, 1))
Else
    str = "select * from PATOOLSLIB.ta101 where case_no=?"
    Set SEL_TA101_QR = rdb.CreateQuery("", str)
    SEL_TA101_QR(0) = Trim(MSFlexGrid1.TextMatrix(LP, 2))
End If

age = ""
SEX = ""
Set SEL_TA101_RS = SEL_TA101_QR.OpenResultset(2, 3)
If Len(Trim(MSFlexGrid1.TextMatrix(LP, 2))) = 6 Then
'age = DateDiff("YYYY", SEL_TA101_RS("AGE"), CStr(DATE11))
age = SEL_TA101_RS("AGE")
SEX = Trim(SEL_TA101_RS("SEX"))
pat_catg = "ST"

Else
age = DateDiff("YYYY", SEL_TA101_RS("BIRTH_DATE"), CStr(DATE11))
SEX = Trim(SEL_TA101_RS("SEX"))
pat_catg = Trim(SEL_TA101_RS("PAT_CATG"))

End If
Dim dmg1 As String
If Len(Trim(MSFlexGrid1.TextMatrix(LP, 2))) = 6 Then
dmg1 = ""
Else
dmg1 = Trim(SEL_TA101_RS("UD_NAME"))
End If



INSTEMPREC_QR(0) = srno
INSTEMPREC_QR(1) = MSFlexGrid1.TextMatrix(LP, 1) 'REQNO
INSTEMPREC_QR(2) = Mid(Trim(MSFlexGrid1.TextMatrix(LP, 1)), 4, 2) 'MODALITYCD
INSTEMPREC_QR(3) = MSFlexGrid1.TextMatrix(LP, 2) 'CASENO
INSTEMPREC_QR(4) = MSFlexGrid1.TextMatrix(LP, 3) 'NAME
INSTEMPREC_QR(5) = SEX 'SEX
INSTEMPREC_QR(6) = age 'AGE
INSTEMPREC_QR(7) = "" 'IPADD (COLLECTION)
INSTEMPREC_QR(8) = MSFlexGrid1.TextMatrix(LP, 10) 'PAT_STAT
INSTEMPREC_QR(9) = pat_catg 'PAT_CATG
INSTEMPREC_QR(10) = cmb_location.Text 'LOCATION
INSTEMPREC_QR(11) = GET_REVDATE(MSFlexGrid1.TextMatrix(LP, 5)) 'COLLDATE
INSTEMPREC_QR(12) = MSFlexGrid1.TextMatrix(LP, 6) 'COLLTIME
INSTEMPREC_QR(13) = "test"   'SER_CODE
INSTEMPREC_QR(14) = MSFlexGrid1.TextMatrix(LP, 9) 'SER_DET
INSTEMPREC_QR(15) = ""     'SENT FLAG
INSTEMPREC_QR(16) = Trim(logonfrm.usrid) 'USERID
INSTEMPREC_QR(17) = MSFlexGrid1.TextMatrix(LP, 7) 'FLD1 (PATHNO)
INSTEMPREC_QR(18) = GET_REVDATE(DATE11)  'RECDDATE
INSTEMPREC_QR(19) = time1  'RECTIME
INSTEMPREC_QR(20) = dmg1 'RECDFLAG
INSTEMPREC_QR(21) = Winsock1.LocalIP  'RECDIAPP
INSTEMPREC_QR(22) = GET_REVDATE(DATE11)  'PR DATE
INSTEMPREC_QR(23) = time1  'PR TIME
INSTEMPREC_QR(24) = MSFlexGrid1.TextMatrix(LP, 8) 'OPPROCE
INSTEMPREC_QR.Execute
abc:
Next
End Sub

Function NOTPRINTED() As Boolean
TXTREQNO.Visible = True
Dim I12 As Integer
Dim TREC_QR As rdoQuery
Dim TREC_RS As rdoResultset
str = "SELECT * FROM Dislib.TEMPRECEIV WHERE SENTDATE=CURRENT DATE AND SENTFLAG<>'P' AND REQNO IN (SELECT REQNNO FROM RISLIB.RREQUISITN WHERE REQNSTAT = 'C' or reqnstat ='E')" 'DREQSTAFFD
    'Set TREC_QR = rdb.CreateQuery("", str)
    Set TREC_QR = rdb.CreateQuery("", str)
    Set TREC_RS = TREC_QR.OpenResultset(2, 3)
    If TREC_RS.EOF = False Then
    cmb_location.Text = "HAEMATOPATHOLOGY LABORATORY$UZZ "
    For I12 = 0 To TREC_RS.RowCount - 1
    
    np = True
    TXTREQNO.Text = Mid(Trim(TREC_RS("REQNO")), 4) & "A"
    'READDATA
    NOTPRINTED = True
    
    Next
    Else
    np = False
    NOTPRINTED = False
    End If
    
End Function
Private Sub Form_Activate()
TXTREQNO.SetFocus
End Sub

Private Sub Form_Load()




str = "SELECT DEPT_CODE,MODLITYDES,MODALITYCD FROM RISLIB.RMODALITY WHERE MODALITYCD=?"
Set SELDEPT_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set GETCASE_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM PATOOLSLIB.TA101 WHERE CASE_NO=?"
Set GETNAME_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=?"
Set GETNAMESTAFF_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELREQNSTAT_QR = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RTIMELOG SET SAMRECDATE=?, SAMRECTIME=? WHERE REQNNO=? "
Set UPDRTIME_QR = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT=?,LOCATION=?,APNTDT=? WHERE REQNNO=?"
Set UPDRREQ_QR = rdb.CreateQuery("", str)

'***** Shraddha 20/12/2011 *****
str = "INSERT INTO INTERFACE.BINTERFACE(CASENO,REQNO,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?)"
Set INSBINTERFACE_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM INTERFACE.BINTERFACE WHERE REQNO=?"
Set SEL_BINTER_QR = rdb.CreateQuery("", str)
'*******************************


'***** Shraddha 24/12/09 *****
str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SEL_PathDreq_QR = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT='C' WHERE REQNNO=?"
Set UPD_RREQNSTAT_QR = rdb.CreateQuery("", str)
'*****************************

str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
Set SELMA101_QR = rdb.CreateQuery("", str)


str = "SELECT * FROM Dislib.TEMPCOLLCT WHERE REQNO=? AND SER_CODE=?"
Set SEL_TEMPCOLLCTR_QR = rdb.CreateQuery("", str)

'***** 22/09/11 *****
str = "SELECT * FROM DISLIB.DTMNO WHERE CASENO=?"
Set SELDTMNO_QR = rdb.CreateQuery("", str)
'*********************

'***** AUTOGENERATION OF THE PATHNO (SERELOGY) *****
'Shraddha Pandit 27/11/08
'str = "SELECT * FROM kalpana.DAUTOPATH WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
str = "SELECT * FROM dislib.DAUTOPATH WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
Set SELSEPATHNO_QR = rdb.CreateQuery("", str)

'***** AUTOGENERATION OF THE PATHNO. Haemato ***** Shraddha 12/10/09
str = "SELECT * FROM DISLIB.DAUTOPATH WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
Set SELDAUTOPATH_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DPATHYEAR WHERE AYEAR=?"
Set DPATHYEAR_QR = rdb.CreateQuery("", str)
'***************************************************************

'***** Haemato Interfaces Shraddha 14/10/09 *****
str = "INSERT INTO INTERFACE.HINTERFACE(CASENO,REQNO,URGENT,IFLAG,LABNO,TDATE) VALUES(?,?,?,?,?,?)"
Set INS_IHINTERFACE_QR = rdb.CreateQuery("", str)
'Added by Shraddha on 12/10/10
str = "INSERT INTO DISLIB.HINTERFACE(CASENO,REQNO,URGENT,IFLAG,LABNO,TDATE) VALUES(?,?,?,?,?,?)"
Set INSHINTERFACE_QR = rdb.CreateQuery("", str)

str = "INSERT INTO INTERFACE.CINTERFACE(CASENO,REQNO,TESTCODE,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?,?)"
Set INSCINTERFACE_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DHEMCYTO1 WHERE REQNO = ?"
Set SELDHEMCYTO_QR = rdb.CreateQuery("", str)

str = "select * from dislib.dhlprosub where hlprocode = ?"
Set seldhlprosub_qr = rdb.CreateQuery("", str)

str = "INSERT INTO INTERFACE.CINTERFACE(CASENO,REQNO,LABNO,TESTCODE,URGENT,IFLAG,TTIME,TDATE) VALUES(?,?,?,?,?,?,?,?)"
Set INS_CINTERFACE_QR = rdb.CreateQuery("", str)
'***************************************************************

'***** Biochem Shraddha 16/10/09 *****
str = "SELECT * FROM DISLIB.DBCPROSUB WHERE BCPROCODE=?"
Set SELDBCPROSUB_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.INTEREXPCD WHERE excpcode =?"
Set interexpcd_qr = rdb.CreateQuery("", str)

str = "INSERT INTO INTERFACE.DINTERFACE(CASENO,REQNO,TESTCODE,URGENT,IFLAG,LABNO,FLD2) VALUES(?,?,?,?,?,?,?)"
Set INSDINTERFACE_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
Set seldmulproce_qr1 = rdb.CreateQuery("", str)

'***** Shraddha 24/08/2011 *****
str = "SELECT * FROM DISLIB.TMTESTCODE WHERE SER_CODE=?"
Set sel_TMTESCODE_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.TMCATEGORY WHERE TYPE=? and STATUS=?"
Set sel_TMCATEGORY_qr = rdb.CreateQuery("", str)
' '***** Shweta 26/04/2012 *****
''str = "SELECT * FROM DISLIB.TMCATEGORY WHERE STATUS='RF'"
'str = "SELECT * FROM DISLIB.TMAUTOPATH WHERE STATUS='RF'"
'Set sel_TMCATEGORYRF_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.TMCATEGORY WHERE TYPE=?"
Set sel_TMCATEGORYANY_qr = rdb.CreateQuery("", str)

str = "UPDATE DISLIB.TMCATEGORY SET PATHNO=? WHERE TYPE=? AND STATUS=?"
Set UPD_TMCATEGORY_qr = rdb.CreateQuery("", str)

str = "UPDATE DISLIB.TMCATEGORY SET PATHNO=? WHERE STATUS='RF'"
Set UPD_TMCATEGORYRF_qr = rdb.CreateQuery("", str)

str = "UPDATE DISLIB.TMCATEGORY SET PATHNO=? WHERE TYPE=?"
Set UPD_TMCATEGORYANY_qr = rdb.CreateQuery("", str)
'*******************************


 '***** Shweta 26/04/2012 *****
'str = "SELECT * FROM DISLIB.TMCATEGORY WHERE STATUS='RF'"
str = "SELECT * FROM DISLIB.TMAUTOPATH WHERE STATUS='RF'"
Set sel_TMCATEGORYRF_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.TMAUTOPATH WHERE TYPE=? and STATUS=?"
Set sel_TMAUTOPATH_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.TMAUTOPATH WHERE TYPE=?"
Set UPD_TMAUTOPATHANY_qr = rdb.CreateQuery("", str)

str = "UPDATE DISLIB.TMAUTOPATH SET PATHNO=? WHERE TYPE=? AND STATUS=?"
Set UPD_TMAUTOPATH_qr = rdb.CreateQuery("", str)

str = "UPDATE DISLIB.TMAUTOPATH SET PATHNO=? WHERE STATUS='RF'"
Set UPD_TMCATEGORYRF_qr = rdb.CreateQuery("", str)

str = "UPDATE DISLIB.TMAUTOPATH SET PATHNO=? WHERE TYPE=? AND STATUS='ANY'"
Set UPD_TMAUTOPATHANY_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.TMAUTOPATH WHERE TYPE=?"
Set Sel_TMAUTOPATHANY_QR = rdb.CreateQuery("", str)

'*******************************

   
str = "SELECT * FROM DISLIB.DELPROSUB WHERE PROCODE=?"
Set seldelprosub_qr = rdb.CreateQuery("", str)

str = "SELECT SAMPLECD FROM DISLIB.SAMPLECODE WHERE SERCODE=?"
Set SAMPLECODE_QR1 = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DREQUISITN WHERE PATHNO=? AND  REQNO LIKE ?"
'Commented and added by Shraddha on 19/07/10
'str = "SELECT * FROM DISLIB.DREQLF_R WHERE PATHNO=? AND  REQNO LIKE ?"
Set SELDREQUISITNQR = rdb.CreateQuery("", str)

str = "INSERT INTO INTERFACE.DINTERFACE(CASENO,REQNO,LABNO,TESTCODE,IFLAG,TTIME,TDATE,fld2) VALUES(?,?,?,?,?,?,?,?)"
Set insdinterface_qr1 = rdb.CreateQuery("", str)


'***** Shraddha 05/11/09 *****
'Dim UPD_RTIME_QR As rdoQuery
str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO = ?"
Set SEL_RREQUP_QR = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RTIMELOG SET SAMCOLDATE=?, SAMCOLTIME=?, SAMRECDATE=?, SAMRECTIME=? WHERE REQNNO=? "
Set upd_rtime_qr = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RTIMELOG SET SAMCOLTIME=? WHERE REQNNO=? "
Set UPD_mTIME_QR = rdb.CreateQuery("", str)

'SAMCOLDATE  SAMCOLTIME  SAMRECDATE  SAMRECTIME
'***************************************

str = "UPDATE DISLIB.DAUTOPATH SET PATHNO=? WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
Set UPDDAUTOPATH_QR = rdb.CreateQuery("", str)


str = "INSERT INTO DISLIB.DAUTOPATH (DEPT_CODE,MODALITYCD,PATHYEAR,PATHYEAR1,PATHNO) VALUES (?,?,?,?,?)"
Set INSDAUTOPATH_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM PATOOLSLIB.MA106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='VZZ') "
Set SELMA106_QR = rdb.CreateQuery("", str)
Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)

str = "UPDATE  RISLIB.RREQUISITN SET EXAMDT=?, REQNSTAT=?,EXAMINATN=?,LOCATION=?,APNTDT=? WHERE REQNNO=?"
Set updrrequisitn_qr = rdb.CreateQuery("", str)

'str = "SELECT SAMRECDATE,SAMRECTIME FROM RISLIB.RTIMELOG WHERE REQNNO=?" 'Commented and added on 05/11/09 by Shraddha
str = "SELECT REQNNO, SAMRECDATE, SAMRECTIME, SAMRECTIME -'01:15:27' FROM RISLIB.RTIMELOG WHERE REQNNO=?"
Set SELTIME_QR = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RTIMELOG SET SAMRECDATE=?,SAMRECTIME=?,CASENO=?,REQNDATE=?, REQNTIME=? WHERE REQNNO=? "
Set UPDTIME_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELREQNOQR = rdb.CreateQuery("", str)

str = "UPDATE DISLIB.DREQUISITN SET PATHNO=? WHERE REQNO=?"
Set upddrequisitn = rdb.CreateQuery("", str)

'***** Added by Shraddha on 04/08/2011 *****
str = "UPDATE DISLIB.TEMPCOLLCT SET FLD1=?,RECDIPAD=? WHERE REQNO=?"
Set upd_TEMMPCOLL_QR = rdb.CreateQuery("", str)


str = "UPDATE DISLIB.TEMPCOLLCT SET SENTFLAG='P' WHERE REQNO=?"
Set upd_TCOLL_Print_QR = rdb.CreateQuery("", str)

'***** Shraddha 09/11/2011 *****
str = "UPDATE DISLIB.TEMPCOLLCT SET SENTFLAG='P' WHERE RECDIPAD=?"
Set upd_TCOLL_P_QR = rdb.CreateQuery("", str)
'*******************************

'***** Shraddha 19/10/2011 *****
str = "UPDATE DISLIB.TEMPCOLLCT SET SENTFLAG='' WHERE REQNO=?"
Set upd_TCOLL_blank_QR = rdb.CreateQuery("", str)
'*******************************

str = "UPDATE DISLIB.TEMPCOLLCT SET RECDIPAD=? WHERE REQNO=?"
Set upd_TCOLL_RCVIP_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.TEMPCOLLCT WHERE REQNO=?"
Set SEL_TEMPCOLL_QR = rdb.CreateQuery("", str)

str = "INSERT INTO DISLIB.TEMPCOLLCT (SRNO,REQNO,MODALITYCD,CASENO,NAME,SEX,AGE,RECDIPAD,PATNSTAT,PAT_CATG,LOCATION,COLLDATE,COLLTIME,SENTDATE,SENTTIME,SER_CODE,SER_DET,SENTFLAG,CFLAG,USERID,FLD1) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
Set INS_TEMPCOLLCT = rdb.CreateQuery("", str)
'*******************************************

str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
Set seldmulproce_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DBCPROSUB WHERE BCPROCODE=?"
Set SELDBCPROSUB_QR1 = rdb.CreateQuery("", str)
 
str = "SELECT SAMPLECD FROM DISLIB.SAMPLECODE WHERE SERCODE=?"
Set SELSAMPLECODE_QR = rdb.CreateQuery("", str)

str = "INSERT INTO INTERFACE.mINTERFACE(CASENO,REQNO,labno,TESTCODE,IFLAG,PATHNO) VALUES(?,?,?,?,?,?)"
Set ins_minterface = rdb.CreateQuery("", str)



'-------ADDED BY SHWETA FOR MICRO AUTO PATH NO.

str = "UPDATE DISLIB.MICROPATH SET PATHNO=?,YEAR1=RIGHT(YEAR(CURRENT DATE),2) ,MONTH1=MONTH(CURRENT DATE) WHERE MODALITY=? AND STATUS=? AND TYPE = ?"
Set UPD_MICROAUTOPATH_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.MICROPATH  WHERE MODALITY=?  AND TYPE = ?"
Set sel_MICROAUTOPATH_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.TMTESTCODE WHERE SER_CODE=?"
Set Sel_MICROTESTCODE_QR = rdb.CreateQuery("", str)

'------------------------

'----------------------------------------------------

txt_pathno.Visible = False 'Shraddha 15/10/09

flg = False

'----------------------
'TXTREQNO.SetFocus
'cmb_location.AddItem ("GEN.OPD") 'Shraddha Pandit 27/11/08
'cmb_location.AddItem ("PVT.OPD")
'cmb_location.ListIndex = 0 'Shraddha Pandit 27/11/08
'cmb_location.Enabled = False 'Commented by Shraddha Pandit 27/11/08
'TXTREQNO.SetFocus
'----- Added by Shraddha Pandit 27/11/08 -----
If SELMA106_RS.RowCount > 0 Then
    'cmb_dept.clear
    cmb_location.clear
    While Not SELMA106_RS.EOF
        cmb_location.AddItem Trim(SELMA106_RS("DEPT_NAME")) & "$" & SELMA106_RS("DEPT_CODE")
   ' If SELMA106_RS("DEPT_CODE") = USRDEPT Then
    'cmb_location.Text = Trim(SELMA106_RS("DEPT_NAME")) & "$" & SELMA106_RS("DEPT_CODE")
'    Else
 '   End If
        
        SELMA106_RS.MoveNext
    Wend
End If
'

'----------------------------------------------
DATETIME
txt_date.Text = DATE11

ipadd = Winsock1.LocalIP
'NOTPRINTED
EnableCloseButton Me.hwnd, False
End Sub

Private Sub Timer1_Timer()
DATETIME
txt_time.Text = time1





End Sub


Private Sub txt_pathno_LostFocus()
If txt_pathno.Visible = True And txt_pathno.Text = "" Then
    MsgBox "Please enter MICRO No"
    txt_pathno.SetFocus
End If
End Sub

Private Sub TXTREQNO_Change()
pend = False
'MsgBox "Requisition ..... ", vbOKOnly
If Len(Trim(TXTREQNO.Text)) = 11 Then
    
    DoEvents
    Dim dept_s As String, mod_s As String, dept_s1 As String
    mod_s = Left(Trim(TXTREQNO.Text), 2)
    Dim rs_deptcode As rdoResultset
    Set rs_deptcode = rdb.OpenResultset("SELECT MODALITYCD,DEPT_CODE FROM RISLIB.RMODALITY WHERE MODALITYCD='" & mod_s & "' ", 1, 1)
    If Not rs_deptcode.EOF Then
    dept_s = Trim(rs_deptcode("DEPT_CODE"))
    End If
    
    
     Set rs_deptcode = rdb.OpenResultset("SELECT DEPT_NAME,DEPT_CODE FROM PATOOLSLIB.MA106 WHERE DEPT_CODE='" & dept_s & "' ", 1, 1)
    If Not rs_deptcode.EOF Then
    dept_s1 = Trim(rs_deptcode("DEPT_NAME"))
    End If
    
    
    If dept_s <> "" Then
    cmb_location.AddItem (dept_s1 & "$" & dept_s)
    cmb_location.Text = dept_s1 & "$" & dept_s
    End If
    
    READDATA
    
End If

End Sub

Public Sub READDATA()
pend = False
'On Error GoTo MYERR

    I = InStr(Trim(cmb_location.Text), "$")


dept = Trim(Mid(cmb_location.Text, I + 1, 4))

If UCase(Trim(Left(TXTREQNO.Text, 2))) = "HL" Or UCase(Trim(Left(TXTREQNO.Text, 2))) = "BC" Then
    If UCase(Trim(Right(TXTREQNO.Text, 1))) = "E" Then
    dept = "FED"
    End If
End If


Dim K As Integer
Dim REC1 As Double
Dim rqno As String
rqno = Trim(Left(TXTREQNO.Text, 10))
    
DATETIME
COLLDATE = DATE11
COLLTIME = time1
  
REQNOB = Trim(TXTREQNO.Text)

REQ111 = Trim(Mid(REQNOB, 3, 8))
    
MODALITYB = Trim(UCase(Left(REQNOB, 2)))
        
REQNOBARCODE = MODALITYB & REQ111

'If Left(REQNOB, 3) <> dept Then
'MsgBox ("This requisition number does not belongs to " & Trim(Left(cmb_location.Text, InStr(cmb_location.Text, "$"))))
'End If

''' COOMENTED AS PER DISCUSSED WITH SWETA MAM BY RAHUL AT SANGRUR ON 04.04.2016  FOR TM MOADILITY COULDN'T RECEIVE..
'''FROM HERE
'If MODALITYB = "TM" Then
'    'Exit the process for  other than FEC staff
'    str = "SELECT USR_ID FROM DISLIB.DSECURITY, DISLIB.REPWRITERS "
'    str = str + "Where USER_ID = USR_ID And USRTYPE = '8' AND DEPT_CODE='FEC'"
'
'    Set REPWTRN_QR = rdb.CreateQuery("", str)
'    Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
'
'    While Not REPWTRNRS.EOF
'        If Trim(REPWTRNRS("USR_ID")) = Trim(UID) Then
'            GoTo MM
'        End If
'        REPWTRNRS.MoveNext
'    Wend
'
'    Exit Sub
'    '**************************
'End If
''' TO TILL HERE ... COOMENTED AS PER DISCUSSED WITH SWETA MAM BY RAHUL AT SANGRUR ON 04.04.2016  FOR TM MOADILITY COULDN'T RECEIVE..
MM:

'************ Validation to check duplicate req no ********
 Dim u
 
 u = 1

While (u < MSFlexGrid1.Rows)
       
    If REQNOBARCODE = Trim(Mid(MSFlexGrid1.TextMatrix(u, 1), 4, 10)) Then
        MsgBox "Details of this requisition already displayed."
        Exit Sub
    Else
        u = u + 1
    End If
Wend
'***********************************************************
Dim RS_rrEQTN  As rdoResultset
 Dim RS_rrtime  As rdoResultset
 Dim RS_INSrrtime  As rdoQuery
  JSTR1 = "SELECT *  FROM RISLIB.RTIMELOG WHERE REQNNO='" & Trim(dept) & Trim(rqno) & "'"
  Dim INSRTIME As String
      Set RS_rrtime = rdb.OpenResultset(JSTR1, 1, 1)
      If RS_rrtime.EOF Then
      
         JSTR1 = "SELECT *  FROM RISLIB.RREQUISITN WHERE REQNNO='" & Trim(dept) & Trim(rqno) & "'"
           
                Set RS_rrEQTN = rdb.OpenResultset(JSTR1, 1, 1)
                If RS_rrEQTN.EOF = False Then
               INSRTIME = "INSERT INTO RISLIB.RTIMELOG(REQNNO,CASENO,REQNDATE,REQNTIME)VALUES(?,?,?,?)"
               Set RS_INSrrtime = rdb.CreateQuery("", INSRTIME)
                RS_INSrrtime(0) = RS_rrEQTN("REQNNO")
                RS_INSrrtime(1) = RS_rrEQTN("CASENO")
                RS_INSrrtime(2) = RS_rrEQTN("REQNDT")
                RS_INSrrtime(3) = RS_rrEQTN("REQNTM")
                RS_INSrrtime.Execute
           End If
      End If



If Trim(USRDEPT) = "FEC" Then
 
    JSTR = "SELECT REQNO,PRO_CODE FROM DISLIB.DMULPROCE WHERE REQNO='" & Trim(dept) & Trim(rqno) & "' AND PRO_CODE IN(SELECT SER_CODE FROM RISLIB.RINVESTIGN WHERE REGION='A')"
    
    Dim RS_JDMULRINEST  As rdoResultset
      Set RS_JDMULRINEST = rdb.OpenResultset(JSTR, 1, 1)
      If RS_JDMULRINEST.EOF = False Then
            MsgBox ("This Test done at Emergency Lab Department so Requisition will not receive")
            TXTREQNO.Text = ""
            TXTREQNO.SetFocus
            Exit Sub
      End If
      
End If

If Trim(USRDEPT) = "FED" Then
    
    JSTR1 = "SELECT REQNO,PRO_CODE FROM DISLIB.DMULPROCE WHERE REQNO='" & Trim(dept) & Trim(rqno) & "' AND PRO_CODE IN(SELECT SER_CODE FROM RISLIB.RINVESTIGN WHERE REGION='B')"
    
    Dim RS_JDMULRINEST1  As rdoResultset
      Set RS_JDMULRINEST1 = rdb.OpenResultset(JSTR1, 1, 1)
      If RS_JDMULRINEST1.EOF = False Then
            MsgBox ("This Test done at Emergency Lab Department so Requisition will not receive")
            TXTREQNO.Text = ""
            TXTREQNO.SetFocus
            Exit Sub
      End If
      
End If
       
SELDEPT_QR(0) = MODALITYB
Set SELDEPT_RS = SELDEPT_QR.OpenResultset(1, 3)
If SELDEPT_RS.EOF = False Then
DEPTB = Trim(SELDEPT_RS("DEPT_CODE"))
Else
MsgBox "Please Enter valid Requisition No."
TXTREQNO.Text = ""
TXTREQNO.SetFocus
'TXTREQNO.SelStart = 0
'TXTREQNO.SelLength = Len(TXTREQNO.Text)
Exit Sub
End If
If Trim(dept) = "FED" And (DEPTB = "UZZ" Or DEPTB = "FEC") Then
Else
If Trim(dept) <> DEPTB Then
    MsgBox "Details can't be displayed, Since" + vbCrLf + "it doesn't belong to selected department"
    TXTREQNO.Text = ""
    TXTREQNO.SetFocus
    Exit Sub
End If
End If
'get reqno
If Trim(dept) = "FED" Then
REQNOB = dept & REQNOBARCODE
Else
REQNOB = DEPTB & REQNOBARCODE
End If

    
SELREQNSTAT_QR(0) = REQNOB
Set SELREQNSTAT_RS = SELREQNSTAT_QR.OpenResultset(1, 3)



'****** Shraddha 04/01/2011 *****
'Z CONDITION IS ADDED BY SHWETA 30-08-2012
If IsNull(SELREQNSTAT_RS("REQNSTAT")) = True Then
MsgBox "Payment not Done."
Exit Sub
End If


If Trim(SELREQNSTAT_RS("REQNSTAT")) = "M" Or Trim(SELREQNSTAT_RS("REQNSTAT")) = "Z" Or Trim(SELREQNSTAT_RS("REQNSTAT")) = "T" Or Trim(SELREQNSTAT_RS("REQNSTAT")) = "F" Then
    MsgBox "Please check this case in Query Status"
    TXTREQNO.Text = ""
    'MSFlexGrid1.clear COMMENTED BY CHARULATA ON 28-02-2012 FOR NOT CLEARING THE GRID
    Exit Sub
End If

If Trim(SELREQNSTAT_RS("REQNSTAT")) = "E" Then   'Shraddha Pandit 26/11/08''***''$
    DoEvents
    MsgBox "Sample already received....! ", vbOKOnly
    
    TXTREQNO.Text = ""
    TXTREQNO.SetFocus
    'TXTREQNO.SelText = True
    Exit Sub
    DoEvents
'ElseIf Trim(SELREQNSTAT_RS("REQNSTAT")) = "P" Then
'    DoEvents
'    MsgBox "Sample Not Collected....! ", vbOKOnly
'    TXTREQNO.Text = ""
'    DoEvents
'    Exit Sub
ElseIf Trim(SELREQNSTAT_RS("REQNSTAT")) = "R" Or Trim(SELREQNSTAT_RS("REQNSTAT")) = "C" Or Trim(SELREQNSTAT_RS("REQNSTAT")) = "P" Then 'Added on 14/10/09''***''$
'ElseIf Trim(SELREQNSTAT_RS("REQNSTAT")) = "C" Then 'Commented and added on 15/11/2011 by Shraddha
    
    PAT_NAME = ""
    
    GETCASE_QR(0) = REQNOB
    Set GETCASE_RS = GETCASE_QR.OpenResultset(1, 3)
     
    If Len(Trim(GETCASE_RS("CASENO"))) = 6 Then
        GETNAMESTAFF_QR(0) = REQNOB
        Set GETNAMESTAFF_RS = GETNAMESTAFF_QR.OpenResultset(1, 3)
        PAT_NAME = GETNAMESTAFF_RS("PNAME")
    Else
        GETNAME_QR(0) = GETCASE_RS("CASENO")
        Set GETNAME_RS = GETNAME_QR.OpenResultset(1, 3)
        PAT_NAME = Trim(GETNAME_RS("TAG")) & " " & Trim(GETNAME_RS("FIRST_NAME")) & " " & Trim(GETNAME_RS("MID_NAME")) & " " & Trim(GETNAME_RS("LAST_NAME"))
    End If
            
    With MSFlexGrid1
        .TextMatrix(0, 0) = "S. No."
        .TextMatrix(0, 1) = "Requisition No."
        .TextMatrix(0, 2) = "Case No."
        .TextMatrix(0, 3) = "Name"
        .TextMatrix(0, 4) = "Modality"
        .TextMatrix(0, 5) = "Receival Date"
        .TextMatrix(0, 6) = "Receival Time"
        .TextMatrix(0, 7) = "MICRO No."          '''''''''''''REPLACE PATH NO WITH MICRO NO BY VARINDER
'        If DEPTB = "UZZ" And MODALITYB = "HL" Then
'        .TextMatrix(0, 9) = "Procedure"
'        .TextMatrix(0, 10) = "Status"
'        End If
        .Rows = .Rows + 1
        .row = .Rows - 2
        
          If Trim(dept) = "FED" Then
          I12 = 0
          For kp = 0 To .Cols - 1
            .col = kp
            .CellBackColor = vbCyan
       
            Next
           End If
        .TextMatrix(.row, 0) = .row
        .TextMatrix(.row, 1) = REQNOB
        .TextMatrix(.row, 2) = GETCASE_RS("CASENO")
        .TextMatrix(.row, 3) = Trim(PAT_NAME)
        .TextMatrix(.row, 4) = Trim(SELDEPT_RS("MODLITYDES"))
        
        
        
        '***** Shraddha 05/11/09 *****
        If Trim(SELREQNSTAT_RS("REQNSTAT")) = "P" And (Trim(SELREQNSTAT_RS("DEPT_CODE")) = "FEC" Or Trim(SELREQNSTAT_RS("DEPT_CODE")) = "FED") Then   ''$$$
            'SELTIME_QR
            'SELTIME_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
            'Set SELTIME_RS = SELTIME_QR.OpenResultset(1, 3)
            
            
            '***** 06/11/09 *****
            'UPDATE RISLIB.RTIMELOG SET SAMCOLDATE=?, SAMCOLTIME=?, SAMRECDATE=?, SAMRECTIME=? WHERE REQNNO=?
            upd_rtime_qr(0) = COLLDATE
            upd_rtime_qr(1) = COLLTIME
            
            upd_rtime_qr(2) = COLLDATE
            upd_rtime_qr(3) = COLLTIME
            upd_rtime_qr(4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
            upd_rtime_qr.Execute
            '********************
            SELTIME_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
            Set SELTIME_RS = SELTIME_QR.OpenResultset(1, 3)
            
            If Not SELTIME_RS.EOF Then
                If Trim(SELTIME_RS(1)) <> "" Then
                    .TextMatrix(.row, 5) = Trim(SELTIME_RS(1))
                Else
                    .TextMatrix(.row, 5) = COLLDATE
                End If
                Dim h, m, S, tm_len
                tm_len = Len(Trim(SELTIME_RS(3)))
                If tm_len = 6 Then
                    h = Mid(Trim(SELTIME_RS(3)), 1, 2)
                    m = Mid(Trim(SELTIME_RS(3)), 3, 2)
                    S = Mid(Trim(SELTIME_RS(3)), 5, 2)
                ElseIf tm_len = 5 Then
                    h = Mid(Trim(SELTIME_RS(3)), 1, 1)
                    m = Mid(Trim(SELTIME_RS(3)), 2, 2)
                    S = Mid(Trim(SELTIME_RS(3)), 4, 2)
                End If
                'Trim (SELTIME_RS(3))
                .TextMatrix(.row, 5) = Trim(SELTIME_RS(1))
                .TextMatrix(.row, 6) = h & ":" & m & ":" & S
                
                '***** 06/11/09 SHRADDHA *****
                'UPD_mTIME_QR(0) = MSFlexGrid1.TextMatrix(.row, 6)
'                UPD_mTIME_QR(0) = h & ":" & m & ":" & S
'                UPD_mTIME_QR(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
'                UPD_mTIME_QR.Execute
                '*****************************
'                Dim DM
'                DM = Trim(SELTIME_RS("SAMRECDATE")) & Trim(SELTIME_RS("SAMRECTIME"))
'                DM = DateAdd("s", -27, DM)
'                DM = DateAdd("m", -15, DM)
'                DM = DateAdd("h", -1, DM)
                '1.15.27
                'SAMRECDATE 'SAMRECTIME
            End If
        Else
        '****************************
            .TextMatrix(.row, 5) = COLLDATE
            .TextMatrix(.row, 6) = COLLTIME
        End If ''' REQNSTAT=P & DEPT=FEC'''$$$
        
        '***** Uncommented by Shraddha on 12/10/09 *****
        .TextMatrix(.row, 7) = "" 'SEAUTONO 'Path / Lab no
        
        '**********ADDED FOR REPEAT SAMPLE
         Dim rs_RepeatSample As rdoResultset
         Set rs_RepeatSample = rdb.OpenResultset("SELECT * FROM dislib.RREQSMSLOG WHERE reqno='" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)) & "' ", 1, 1)
            If Not rs_RepeatSample.EOF Then
            
             'UPDATE RISLIB.RTIMELOG SET SAMCOLDATE=?, SAMCOLTIME=?, SAMRECDATE=?, SAMRECTIME=? WHERE REQNNO=?
            upd_rtime_qr(0) = COLLDATE
            upd_rtime_qr(1) = COLLTIME
            
            upd_rtime_qr(2) = COLLDATE
            upd_rtime_qr(3) = COLLTIME
            upd_rtime_qr(4) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
            upd_rtime_qr.Execute
            '********************
            Dim updrrequ_qr As rdoQuery
            str = "UPDATE  RISLIB.RREQUISITN SET EXAMDT=CURRENT DATE, REQNSTAT='E',APNTDT=CURRENT DATE WHERE REQNNO=?"
            Set updrrequ_qr = rdb.CreateQuery("", str)

            
'                updrrequ_qr(4) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                updrrequ_qr(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                updrrequ_qr.Execute
            
            
            'GETTING PATHNO AND SERVICE DESC
            Dim qr_PATHNO As rdoQuery
            Dim rs_PATHNO As rdoResultset
            Set qr_PATHNO = rdb.CreateQuery("", "select * from dislib.drequisitn where reqno='" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)) & "'")
            Set rs_PATHNO = qr_PATHNO.OpenResultset(1, 3)
                If rs_PATHNO.EOF = False Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = Trim(rs_PATHNO("PATHNO"))
                End If
            
            'SERVICE DESC
            Dim SEL_DMUL_QR1 As rdoQuery
            Dim SEL_DMUL_RS1 As rdoResultset
            
            Dim SEL_MA101_QR1 As rdoQuery
            Dim SEL_MA101_RS1 As rdoResultset
            ROW11 = MSFlexGrid1.row
            str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
            Set SEL_DMUL_QR1 = rdb.CreateQuery("", str)

            SEL_DMUL_QR1(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
            'SELLABNO_QR.Execute
            Set SEL_DMUL_RS1 = SEL_DMUL_QR1.OpenResultset(1, 3)
            TESTDESC = ""
            While Not SEL_DMUL_RS1.EOF
                                          
                str = "SELECT * FROM PATOOLSLIB.MA101 WHERE SER_CODE=?"
                Set SEL_MA101_QR1 = rdb.CreateQuery("", str)
                                          
                SEL_MA101_QR1(0) = Trim(SEL_DMUL_RS1("PRO_CODE"))
                Set SEL_MA101_RS1 = SEL_MA101_QR1.OpenResultset(1, 3)

                'DESCRIP = ""
                'DESCRIP = Mid(Trim(SELDEPTDES_RS("DEPT_NAME")), 1, 4) & "/" & colorcd
                TESTDESC = TESTDESC & "| " & Trim(SEL_MA101_RS1("SER_DET"))
                'TESTDESC = TESTDESC & "| " & vbCrLf & Trim(SEL_MA101_RS("SER_DET"))
                'following line is added by Shweta on 07-05-2012
                ToEnterUnit Trim(SEL_DMUL_RS1("PRO_CODE")), ROW11
                SEL_DMUL_RS1.MoveNext
            Wend
    
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 8) = TESTDESC
            
            
            
            Exit Sub
            End If
        
        
        '**********************************
        
        If (DEPTB = "UZZ" And MODALITYB = "HL" And dept = "UZZ") Then
            SPAUTONO
            
'            SEL_PathDreq_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
'            Set SEL_PathDreq_RS = SEL_PathDreq_QR.OpenResultset(1, 3)
'            If IsNull(Trim(SEL_PathDreq_RS("OPPROCEDUR"))) = False Or Trim(Trim(SEL_PathDreq_RS("OPPROCEDUR"))) <> "" Then
'            .TextMatrix(.row, 9) = Trim(SEL_PathDreq_RS("OPPROCEDUR"))
'            Else
'            .TextMatrix(.row, 9) = ""
'            End If
'
            
            ElseIf (DEPTB = "UZZ") And MODALITYB = "HL" And dept = "FED" Then
            SPAUTONO_FED
           
            
        End If
        If (DEPTB = "FEC") Then   '*
'            If MSFlexGrid1.row > 1 Then
'            'call function
'                FECPATHNO
'                '***** Shraddha 10/02/10 *****
'                If MSFlexGrid1.TextMatrix(0, 0) <> "" Then
'                FEC_ASSIGN
'                assign
'                End If
'                '*****************************
'            Else
                
            '***** Shraddha 24/08/2011 *****
            
'                txt_pathno.Visible = True
'                txt_pathno.SetFocus 'Shraddha 24/12/09
                
                'Commented and added by Shraddha on 24/08/2011
                
                'Read DMULPROCE, see Test code
                'Check its flag in TMTESTCODE
                'According to its status in TMCATEGORY give pahno
                'Increment and update pathno of TMCATEGORY again
                
                
                If MODALITYB = "BC" And dept = "FEC" Then 'And ((Trim(GETCASE_RS("PAT_CATG")) <> "C") And (Trim(GETCASE_RS("PAT_CATG")) <> "NC")) Then'"**
                    'PRIVATE OUTPATIENT
                    '***** BC C *****
                                
                    
                    
                    If ((Trim(GETCASE_RS("PAT_CATG")) <> "BP") And (Trim(GETCASE_RS("PAT_CATG")) <> "C") And (Trim(GETCASE_RS("PAT_CATG")) <> "NC")) And Trim(GETCASE_RS("PATNSTAT")) = "OP" Then
                        patstat = "POP"
                    ElseIf ((Trim(GETCASE_RS("PAT_CATG")) = "BP") Or (Trim(GETCASE_RS("PAT_CATG")) = "C") Or (Trim(GETCASE_RS("PAT_CATG")) = "NC")) And Trim(GETCASE_RS("PATNSTAT")) = "OP" Then
                        patstat = "GOP"
                    ElseIf Trim(GETCASE_RS("PATNSTAT")) = "IP" Then
                        patstat = "IP"
                       
                    End If
                    If Trim(GETCASE_RS("PATNSTAT")) = "OP" Then
                        If Trim(patstat) = "" Then
                        patstat = "IN"
                        End If
                    End If
                        
                    If Len(Trim(GETCASE_RS("CASENO"))) = 6 Then
                        patstat = "ST"
                    End If
                    
                    If Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 1) = "R" And Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 2, 1) <> "A" Then
                        patstat = "RF"
                    End If

                    SEL_PathDreq_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
                    Set SEL_PathDreq_RS = SEL_PathDreq_QR.OpenResultset(1, 3)
                    
                    If Left(Trim(SEL_PathDreq_RS("ICDS2")), 2) = "RM" Then
                    patstat = "IP"
                    'FOLLOWING NEW ICDS2=I1 ADDED BY SHWETA 04-08-2012
                    ElseIf Trim(SEL_PathDreq_RS("ICDS2")) = "YES" Or Trim(SEL_PathDreq_RS("ICDS2")) = "I1" Or Trim(SEL_PathDreq_RS("SPIHREQNO")) = "U" Then
                                        patstat = "IC"
                    End If
                    
                     If UCase(Left(Trim(GETCASE_RS("CASENO")), 1)) = "Y" Or UCase(Left(Trim(GETCASE_RS("CASENO")), 1)) = "W" Then
                        patstat = "YA"
                    End If
'                    If Not SEL_PathDreq_RS.EOF Then
'                        If Trim(SEL_PathDreq_RS("ICDS2")) = "PR" Then
'                            patstat = "PR"
'                        End If
'                    End If
                    
                    seldmulproce_qr1(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
                    Set seldmulproce_rs1 = seldmulproce_qr1.OpenResultset(1, 3)
                    
                    If Not seldmulproce_rs1.EOF Then '1
                        '1.1 FOR HIAA AND VMA(F989,F897)
                      If seldmulproce_rs1("PRO_CODE") = "F897" Or seldmulproce_rs1("PRO_CODE") = "F898" Then '1.1
                             sel_TMTESCODE_qr(0) = seldmulproce_rs1("PRO_CODE")
                             Set sel_TMTESCODE_rs = sel_TMTESCODE_qr.OpenResultset(1, 3)
                        
                              If Not sel_TMTESCODE_rs.EOF Then '2
                                 ''''''''''
                                 
                                 sel_TMCATEGORYANY_qr(0) = Trim(sel_TMTESCODE_rs("TYPE"))
                             
                             sel_TMCATEGORY_qr(0) = Trim(sel_TMTESCODE_rs("TYPE"))
                            sel_TMCATEGORY_qr(1) = "ANY"
                           
                            Set sel_TMCATEGORY_rs = sel_TMCATEGORY_qr.OpenResultset(1, 3)
                        
                            If Not sel_TMCATEGORY_rs.EOF Then '3
                            
                                tmpathno_VH = sel_TMCATEGORY_rs("PATHNO") + 1
                                
                                If Len(tmpathno_VH) = 1 Then
                                tempno = "000" & tmpathno_VH
                                ElseIf Len(tmpathno_VH) = 2 Then
                                tempno = "00" & tmpathno_VH
                                ElseIf Len(tmpathno_VH) = 3 Then
                                tempno = "0" & tmpathno_VH
                                ElseIf Len(tmpathno_VH) = 4 Then
                                tempno = tmpathno_VH
                                End If
                                If seldmulproce_rs1("PRO_CODE") = "F897" Then
                                ST = "V"
                                ElseIf seldmulproce_rs1("PRO_CODE") = "F898" Then
                                ST = "H"
                                End If
                                
                                 MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = ST & tempno
                                
                                    upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                    upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                    upddrequisitn.Execute
                                    
                                    UPD_TMCATEGORY_qr(0) = tempno
                                    UPD_TMCATEGORY_qr(1) = sel_TMTESCODE_rs("TYPE")
                                    UPD_TMCATEGORY_qr(2) = "ANY"
                                    UPD_TMCATEGORY_qr.Execute
                                        
                                    FEC_ASSIGN
            
                                    assign
                                                              
                               End If
                                 
                                 '''''''''''''''''
        
                                Else
                                MsgBox "TEST CODE NOT FOUND IN TMTESTCODE MASTER"
                             End If
                      Else
                    
                    
                        sel_TMTESCODE_qr(0) = seldmulproce_rs1("PRO_CODE")
                       
                        
                        Set sel_TMTESCODE_rs = sel_TMTESCODE_qr.OpenResultset(1, 3)
                        
                        If Not sel_TMTESCODE_rs.EOF Then '2
                            '''''''added by varinder from sangrur 25072023
                            If Trim(sel_TMTESCODE_rs("TYPE")) = "Y" Or Trim(sel_TMTESCODE_rs("TYPE")) = "V" Or Trim(sel_TMTESCODE_rs("TYPE")) = "W" Or Trim(sel_TMTESCODE_rs("TYPE")) = "F" Then

                                patstat = "ALL"

                            End If
                            '***** 08/11/2011 *****
                            
                            sel_TMCATEGORYANY_qr(0) = Trim(sel_TMTESCODE_rs("TYPE"))
                            'Set sel_TMCATEGORYANY_rs = sel_TMCATEGORYANY_qr.OpenResultset(1, 3)
                            'If Trim(patstat) = "PR" Or Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 1) = "R" Then
                            If (Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 1) = "R" And Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 2, 1) <> "A") Then
                                sel_TMCATEGORYANY_qr(0) = "X"
                                Set sel_TMCATEGORYANY_rs = sel_TMCATEGORYANY_qr.OpenResultset(1, 3)
                            End If
                                Set sel_TMCATEGORYANY_rs = sel_TMCATEGORYANY_qr.OpenResultset(1, 3)
                            '***** 26/08/2011 *****
                            If (Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 1) = "R" And Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 2, 1) <> "A") Or Trim(patstat) = "PR" Then '7
                                                       
                                ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
                                 
                                If Left(Trim(sel_TMCATEGORYANY_rs("PATHNO")), 6) = Trim(ST) Then '8
                            
                                     'tmpathno = ST & sel_TMCATEGORY_rs("SERIES")
                                     'Left(GETCASE_RS("PATNSTAT"), 1) = "R"
                                     tmpathno = sel_TMCATEGORYANY_rs("PATHNO") + 1
                                     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = tmpathno
                                     'Format of pathno is like YYMMDD
                                     upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                     upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                     upddrequisitn.Execute
                        
                        
                                        UPD_TMCATEGORYANY_qr(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                     'str = "UPDATE DISLIB.TMCATEGORY SET PATHNO=? WHERE TYPE=? AND STATUS=?"
                                     'If Trim(patstat) = "PR" Or Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 1) = "R" Then
                                     If Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 1) = "R" And Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 2, 1) <> "A" Then
                                        UPD_TMCATEGORYANY_qr(1) = "X"
                                        'Set sel_TMCATEGORYANY_rs = sel_TMCATEGORYANY_qr.OpenResultset(1, 3)
                                        
                                    Else
                                        UPD_TMCATEGORYANY_qr(1) = sel_TMTESCODE_rs("TYPE")
                                    End If
                                     
                                     'UPD_TMCATEGORYANY_qr(1) = sel_TMTESCODE_rs("TYPE")
                                     'UPD_TMCATEGORY_qr(2) = GETCASE_RS("PATNSTAT")
                                     UPD_TMCATEGORYANY_qr.Execute
                        
                                     FEC_ASSIGN

                                     assign

                                 Else '8
                               
                                     If Len(Trim(Trim(sel_TMCATEGORYANY_rs("SERIES")))) = 1 Then
                                         tempno = "000" & Trim(Trim(sel_TMCATEGORYANY_rs("SERIES")))
                                     ElseIf Len(Trim(Trim(sel_TMCATEGORYANY_rs("SERIES")))) = 2 Then
                                         tempno = "00" & Trim(Trim(sel_TMCATEGORYANY_rs("SERIES")))
                                     ElseIf Len(Trim(Trim(sel_TMCATEGORYANY_rs("SERIES")))) = 3 Then
                                         tempno = "0" & Trim(Trim(sel_TMCATEGORYANY_rs("SERIES")))
                                     End If
                
                                     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = ST & tempno
                    
                                     'UPDATE PATHNO IN DREQUISITN
                                     upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                     upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                     upddrequisitn.Execute
                
                                     UPD_TMCATEGORYANY_qr(0) = ST & tempno
                                     
                                     If Trim(patstat) = "PR" Or (Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 1) = "R" And Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 2, 1) <> "A") Then
                                        UPD_TMCATEGORYANY_qr(1) = "X"
                                     Else
                                        UPD_TMCATEGORYANY_qr(1) = sel_TMTESCODE_rs("TYPE")
                                        
                                     End If
                                     
                                     UPD_TMCATEGORYANY_qr.Execute
                     
                                     FEC_ASSIGN

                                     assign
                            
                                End If '8
                            Else
                            '****** 08/11/2011 *****
                            
                            sel_TMCATEGORY_qr(0) = Trim(sel_TMTESCODE_rs("TYPE"))
                            sel_TMCATEGORY_qr(1) = patstat
                           
                            Set sel_TMCATEGORY_rs = sel_TMCATEGORY_qr.OpenResultset(1, 3)
                        
                            If Not sel_TMCATEGORY_rs.EOF Then '3
                                
                                ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
                    
                                If Left(Trim(sel_TMCATEGORY_rs("PATHNO")), 6) = Trim(ST) Then '4
                                                                             
                                    tmpathno = sel_TMCATEGORY_rs("PATHNO") + 1
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = tmpathno
                                    'Format of pathno is like YYMMDD
                                    upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                    upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                    upddrequisitn.Execute
                                    
                                    'str = "UPDATE DISLIB.TMCATEGORY SET PATHNO=? WHERE TYPE=? AND STATUS=?"
                                    UPD_TMCATEGORY_qr(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                    UPD_TMCATEGORY_qr(1) = sel_TMTESCODE_rs("TYPE")
                                    UPD_TMCATEGORY_qr(2) = patstat
                                    UPD_TMCATEGORY_qr.Execute
                                    
                                    FEC_ASSIGN
        
                                    assign
                                    
                                Else '4
                                                       
                                    If Len(Trim(Trim(sel_TMCATEGORY_rs("SERIES")))) = 1 Then
                                        tempno = "000" & Trim(Trim(sel_TMCATEGORY_rs("SERIES")))
                                    ElseIf Len(Trim(Trim(sel_TMCATEGORY_rs("SERIES")))) = 2 Then
                                        tempno = "00" & Trim(Trim(sel_TMCATEGORY_rs("SERIES")))
                                    ElseIf Len(Trim(Trim(sel_TMCATEGORY_rs("SERIES")))) = 3 Then
                                        tempno = "0" & Trim(Trim(sel_TMCATEGORY_rs("SERIES")))
                                    Else
                                        tempno = Trim(Trim(sel_TMCATEGORY_rs("SERIES")))
                                    End If
                                
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = ST & tempno
                                
                                    upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                    upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                    upddrequisitn.Execute
                                    
                                    UPD_TMCATEGORY_qr(0) = ST & tempno
                                    UPD_TMCATEGORY_qr(1) = sel_TMTESCODE_rs("TYPE")
                                    UPD_TMCATEGORY_qr(2) = patstat
                                    UPD_TMCATEGORY_qr.Execute
                                        
                                    FEC_ASSIGN
            
                                    assign
                             
                                End If '4
                            
                            End If '3
                            
                            End If '08/11/2011
                        Else
                            ''HERE WAS CODE
                            ''''''
                        End If '2
                     End If  '1.1 FOR HIAA AND VMA(F989,F897)
                    End If '1  ADDED ON 06/09/2011
                '***** EBCC *****
                ElseIf MODALITYB = "TM" Then '''**
                
                
                If ((Trim(GETCASE_RS("PAT_CATG")) <> "C") And (Trim(GETCASE_RS("PAT_CATG")) <> "NC")) And Trim(GETCASE_RS("PATNSTAT")) = "OP" Then
                        patstat = "POP"
                    ElseIf ((Trim(GETCASE_RS("PAT_CATG")) = "C") Or (Trim(GETCASE_RS("PAT_CATG")) = "NC")) And Trim(GETCASE_RS("PATNSTAT")) = "OP" Then
                        patstat = "GOP"
                    ElseIf Trim(GETCASE_RS("PATNSTAT")) = "IP" Then
                        patstat = "IP"
                       
                    End If
                    If Trim(GETCASE_RS("PATNSTAT")) = "OP" Then
                        If Trim(patstat) = "" Then
                        patstat = "IN"
                        End If
                    End If
                        
                    If Len(Trim(GETCASE_RS("CASENO"))) = 6 Then
                        patstat = "ST"
                    End If
                    
                    If Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 1) = "R" And Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 2, 1) <> "A" Then
                        patstat = "RF"
                    End If
                
                     If UCase(Left(Trim(GETCASE_RS("CASENO")), 1)) = "Y" Or UCase(Left(Trim(GETCASE_RS("CASENO")), 1)) = "W" Then
                        patstat = "YA"
                    End If
                
                
                
                    seldmulproce_qr1(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
                    Set seldmulproce_rs1 = seldmulproce_qr1.OpenResultset(1, 3)
                        Dim TSCD As String
                        TSCD = Trim(seldmulproce_rs1("PRO_CODE"))
                        
                        If TSCD = "F917" Or TSCD = "F053" Or TSCD = "F054" Or TSCD = "F977" Then
                        patstat = "BJP"
                        ElseIf TSCD = "F052" Or TSCD = "F050" Or TSCD = "F051" Or TSCD = "F041" Or TSCD = "F042" Or TSCD = "F044" Or TSCD = "F046" Or TSCD = "F047" Or TSCD = "F048" Then
                        patstat = "EP"
                        End If
                        
                       
                        
                    If Not seldmulproce_rs1.EOF Then '1
                        sel_TMTESCODE_qr(0) = seldmulproce_rs1("PRO_CODE")
                        Set sel_TMTESCODE_rs = sel_TMTESCODE_qr.OpenResultset(1, 3)
                        
                        If Not sel_TMTESCODE_rs.EOF Then '2
                            '***** RF *****
                               If Left(GETCASE_RS("CASENO"), 1) = "R" And Trim(sel_TMTESCODE_rs("TYPE")) = "M" Then '3
                                    'sel_TMCATEGORYRF_qr(0) = Trim(sel_TMTESCODE_rs("TYPE"))
                                    Set sel_TMCATEGORYRF_RS = sel_TMCATEGORYRF_qr.OpenResultset(1, 3)
                                              
                                    ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
                
                                    If Left(Trim(sel_TMCATEGORYRF_RS("PATHNO")), 6) = Trim(ST) Then '4
                             
                                        'tmpathno = ST & sel_TMCATEGORY_rs("SERIES")
                                
                                        tmpathno = sel_TMCATEGORYRF_RS("PATHNO") + 1
                                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = tmpathno
                                        'Format of pathno is like YYMMDD
                                        upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                        upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                        upddrequisitn.Execute
                                        
                                        'str = "UPDATE DISLIB.TMAUTOPATH SET PATHNO=? WHERE TYPE=? AND STATUS=?"
                                        UPD_TMAUTOPATH_qr(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                        UPD_TMAUTOPATH_qr(1) = sel_TMTESCODE_rs("TYPE")
                                        UPD_TMAUTOPATH_qr(2) = "RF"
                                        UPD_TMAUTOPATH_qr.Execute
                                        
                                        FEC_ASSIGN
            
                                        assign
                                    
                                     'End If
                                    Else '4
                                
                                        If Len(Trim(Trim(sel_TMCATEGORYRF_RS("SERIES")))) = 1 Then
                                            tempno = "000" & Trim(Trim(sel_TMCATEGORYRF_RS("SERIES")))
                                        ElseIf Len(Trim(Trim(sel_TMCATEGORYRF_RS("SERIES")))) = 2 Then
                                            tempno = "00" & Trim(Trim(sel_TMCATEGORYRF_RS("SERIES")))
                                        ElseIf Len(Trim(Trim(sel_TMCATEGORYRF_RS("SERIES")))) = 3 Then
                                            tempno = "0" & Trim(Trim(sel_TMCATEGORYRF_RS("SERIES")))
                                            Else
                                        tempno = Trim(Trim(sel_TMCATEGORY_rs("SERIES")))
                                        End If
                            
                                        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = ST & tempno
                            
                                
                                        upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                        upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                        upddrequisitn.Execute
                                    
                                    'str = "UPDATE DISLIB.TMAUTOPATH SET PATHNO=? WHERE TYPE=? AND STATUS=?"
                                        UPD_TMAUTOPATH_qr(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                        UPD_TMAUTOPATH_qr(1) = sel_TMTESCODE_rs("TYPE")
                                        UPD_TMAUTOPATH_qr(2) = "RF"
                                        UPD_TMAUTOPATH_qr.Execute
                                    
                                        FEC_ASSIGN
        
                                        assign
                             
                                    End If '4
                            
                            '*****
                       
                               Else '3
                          '***** E M RF *****
                        If Trim(sel_TMTESCODE_rs("TYPE")) = "N" Then
                        patstat = "EPT"
                        Else
                        End If
                                    sel_TMAUTOPATH_qr(0) = Trim(sel_TMTESCODE_rs("TYPE"))
                                    sel_TMAUTOPATH_qr(1) = patstat 'GETCASE_RS("PATNSTAT")
                                    Set sel_TMAUTOPATH_rs = sel_TMAUTOPATH_qr.OpenResultset(1, 3)
                        
                                    If Not sel_TMAUTOPATH_rs.EOF Then '5                            'Get pathno
                                        '*****
                                        ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
                            
                                        'If Len(Trim(sel_TMCATEGORY_rs("PATHNO"))) > 0 Then
                
                                        If Left(Trim(sel_TMAUTOPATH_rs("PATHNO")), 6) = Trim(ST) Then '6
                             
                                            'tmpathno = ST & sel_TMCATEGORY_rs("SERIES")
                                
                                            tmpathno = sel_TMAUTOPATH_rs("PATHNO") + 1
                                            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = tmpathno
                                            'Format of pathno is like YYMMDD
                                            upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                            upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                            upddrequisitn.Execute
                                    
                                    'str = "UPDATE DISLIB.TMCATEGORY SET PATHNO=? WHERE TYPE=? AND STATUS=?"
                                            UPD_TMAUTOPATH_qr(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                            UPD_TMAUTOPATH_qr(1) = sel_TMTESCODE_rs("TYPE")
                                            UPD_TMAUTOPATH_qr(2) = patstat 'GETCASE_RS("PATNSTAT")
                                            UPD_TMAUTOPATH_qr.Execute
                                    
                                            FEC_ASSIGN
        
                                            assign
                                    
                                            'End If
                                        Else '6
                                
                                            If Len(Trim(Trim(sel_TMAUTOPATH_rs("SERIES")))) = 1 Then
                                                tempno = "000" & Trim(Trim(sel_TMAUTOPATH_rs("SERIES")))
                                            ElseIf Len(Trim(Trim(sel_TMAUTOPATH_rs("SERIES")))) = 2 Then
                                                tempno = "00" & Trim(Trim(sel_TMAUTOPATH_rs("SERIES")))
                                            ElseIf Len(Trim(Trim(sel_TMAUTOPATH_rs("SERIES")))) = 3 Then
                                                tempno = "0" & Trim(Trim(sel_TMAUTOPATH_rs("SERIES")))
                                                Else
                                        tempno = Trim(Trim(sel_TMAUTOPATH_rs("SERIES")))
                                            End If
                            
                                            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = ST & tempno
                            
                                            upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                            upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                            upddrequisitn.Execute
                            
                                            UPD_TMAUTOPATH_qr(0) = ST & tempno
                                            UPD_TMAUTOPATH_qr(1) = sel_TMTESCODE_rs("TYPE")
                                            UPD_TMAUTOPATH_qr(2) = patstat ' GETCASE_RS("PATNSTAT")
                                            UPD_TMAUTOPATH_qr.Execute
                                    
                                            FEC_ASSIGN
        
                                            assign
                             
                                        End If '6
                        
                                    Else '5
                        
                                        Sel_TMAUTOPATHANY_QR(0) = Trim(sel_TMTESCODE_rs("TYPE"))
                                        Set Sel_TMAUTOPATHANY_rs = Sel_TMAUTOPATHANY_QR.OpenResultset(1, 3)
                            
                                        '***** 26/08/2011 *****
                                        'If Trim(sel_TMTESCODE_rs("TYPE")) = "B" Or Trim(sel_TMTESCODE_rs("TYPE")) = "O" Or Trim(sel_TMTESCODE_rs("TYPE")) = "E" Or Trim(sel_TMTESCODE_rs("TYPE")) = "T" Then '7
                                        'Commented and added by Shraddha on 25/11/2011
                                        If Trim(Sel_TMAUTOPATHANY_rs("TYPE")) = "U" Or Trim(sel_TMTESCODE_rs("TYPE")) = "B" Or Trim(sel_TMTESCODE_rs("TYPE")) = "O" Or Trim(sel_TMTESCODE_rs("TYPE")) = "E" Or Trim(sel_TMTESCODE_rs("TYPE")) = "T" Or Trim(sel_TMTESCODE_rs("TYPE")) = "N" Then  '7
                                               ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
                                     
                                               If Left(Trim(Sel_TMAUTOPATHANY_rs("PATHNO")), 6) = Trim(ST) Then '8
                                
                                                    'tmpathno = ST & sel_TMCATEGORY_rs("SERIES")
                                                    'Left(GETCASE_RS("PATNSTAT"), 1) = "R"
                                                    tmpathno = Sel_TMAUTOPATHANY_rs("PATHNO") + 1
                                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = tmpathno
                                                    'Format of pathno is like YYMMDD
                                                    upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                                    upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                                    upddrequisitn.Execute
                                       
                                                    'str = "UPDATE DISLIB.TMCATEGORY SET PATHNO=? WHERE TYPE=? AND STATUS=?"
                                                    UPD_TMAUTOPATHANY_qr(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                                    UPD_TMAUTOPATHANY_qr(1) = sel_TMTESCODE_rs("TYPE")
                                       'UPD_TMCATEGORY_qr(2) = GETCASE_RS("PATNSTAT")
                                                    UPD_TMAUTOPATHANY_qr.Execute
                                       
                                                    FEC_ASSIGN
        
                                                    assign
                                       
                                    'End If
                                                Else '8
                                   
                                                    If Len(Trim(Trim(Sel_TMAUTOPATHANY_rs("SERIES")))) = 1 Then
                                                        tempno = "000" & Trim(Trim(Sel_TMAUTOPATHANY_rs("SERIES")))
                                                    ElseIf Len(Trim(Trim(Sel_TMAUTOPATHANY_rs("SERIES")))) = 2 Then
                                                        tempno = "00" & Trim(Trim(Sel_TMAUTOPATHANY_rs("SERIES")))
                                                    ElseIf Len(Trim(Trim(Sel_TMAUTOPATHANY_rs("SERIES")))) = 3 Then
                                                        tempno = "0" & Trim(Trim(Sel_TMAUTOPATHANY_rs("SERIES")))
                                                    Else
                                                        tempno = Trim(Trim(Sel_TMAUTOPATHANY_rs("SERIES")))
                                                    End If
                               
                                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = ST & tempno
                                   
                                                    'UPDATE PATHNO IN DREQUISITN
                                                    upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                                    upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                                    upddrequisitn.Execute
                               
                                                    UPD_TMAUTOPATHANY_qr(0) = ST & tempno
                                                    UPD_TMAUTOPATHANY_qr(1) = sel_TMTESCODE_rs("TYPE")
                                    
                                                    UPD_TMAUTOPATHANY_qr.Execute
                                    
                                                    FEC_ASSIGN
        
                                                    assign
                                
                                                End If '8
                                                '*****
                            
                                            End If '7
                                   
                                        End If '5
                                End If '3
              
                        '********26/08/11 *****
                End If '2
            End If '1
                
            ElseIf MODALITYB = "EL" Or (MODALITYB = "BC" And dept = "FED") Then
                If ((Trim(GETCASE_RS("PAT_CATG")) <> "C") And (Trim(GETCASE_RS("PAT_CATG")) <> "NC")) And Trim(GETCASE_RS("PATNSTAT")) = "OP" Then
                        patstat = "POP"
                ElseIf ((Trim(GETCASE_RS("PAT_CATG")) = "C") Or (Trim(GETCASE_RS("PAT_CATG")) = "NC")) And Trim(GETCASE_RS("PATNSTAT")) = "OP" Then
                        patstat = "GOP"
                ElseIf Trim(GETCASE_RS("PATNSTAT")) = "IP" Then
                        patstat = "IP"
                End If
                
                If Trim(GETCASE_RS("PATNSTAT")) = "OP" Then
                        patstat = "IN"
                End If
                        
                If Len(Trim(GETCASE_RS("CASENO"))) = 6 Then
                        patstat = "ST"
                End If
                If Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 1) = "R" And Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 2, 1) <> "A" Then
                        patstat = "RF"
                End If
                'COMMENT THIS CODE 19-JAN-2015
                   ' SEL_PathDreq_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
                   ' Set SEL_PathDreq_RS = SEL_PathDreq_QR.OpenResultset(1, 3)
                    
                   ' If Not SEL_PathDreq_RS.EOF Then
                    '    If Trim(SEL_PathDreq_RS("ICDS2")) = "PR" Then
                     '       patstat = "PR"
                     '   End If
                    'End If
                    'TILL HERE
                    seldmulproce_qr1(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
                    Set seldmulproce_rs1 = seldmulproce_qr1.OpenResultset(1, 3)
                    
                    If Not seldmulproce_rs1.EOF Then '1
                        sel_TMTESCODE_qr(0) = seldmulproce_rs1("PRO_CODE")
                        Set sel_TMTESCODE_rs = sel_TMTESCODE_qr.OpenResultset(1, 3)
                        
                        If Not sel_TMTESCODE_rs.EOF Then '2
                            '''''''
                            
                            '***** 08/11/2011 *****
                            
                            sel_TMCATEGORYANY_qr(0) = Trim(sel_TMTESCODE_rs("TYPE"))
                            'Set sel_TMCATEGORYANY_rs = sel_TMCATEGORYANY_qr.OpenResultset(1, 3)
                            
                            If Trim(patstat) = "PR" Or Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 1) = "R" Then
                                sel_TMCATEGORYANY_qr(0) = "X"
                                'Set sel_TMCATEGORYANY_rs = sel_TMCATEGORYANY_qr.OpenResultset(1, 3)
                            End If
                            Set sel_TMCATEGORYANY_rs = sel_TMCATEGORYANY_qr.OpenResultset(1, 3)
                            '***** 26/08/2011 *****
                            If Trim(sel_TMTESCODE_rs("TYPE")) = "X" Or (Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 1) = "R" And Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 2, 1) <> "A") Or Trim(patstat) = "PR" Then  '7
                                                       
                                ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
                                 
                                If Left(Trim(sel_TMCATEGORYANY_rs("PATHNO")), 6) = Trim(ST) Then '8
                            
                                     'tmpathno = ST & sel_TMCATEGORY_rs("SERIES")
                                     'Left(GETCASE_RS("PATNSTAT"), 1) = "R"
                                     tmpathno = sel_TMCATEGORYANY_rs("PATHNO") + 1
                                     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = tmpathno
                                     'Format of pathno is like YYMMDD
                                     upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                     upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                     upddrequisitn.Execute
                        
                                    UPD_TMCATEGORYANY_qr(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                     'str = "UPDATE DISLIB.TMCATEGORY SET PATHNO=? WHERE TYPE=? AND STATUS=?"
                                     If Trim(patstat) = "PR" Or (Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 1) = "R" And Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 2, 1) <> "A") Then
                                        UPD_TMCATEGORYANY_qr(1) = "X"
                                        'Set sel_TMCATEGORYANY_rs = sel_TMCATEGORYANY_qr.OpenResultset(1, 3)
                                        
                                    Else
                                        UPD_TMCATEGORYANY_qr(1) = sel_TMTESCODE_rs("TYPE")
                                    End If
                                     
                                     'UPD_TMCATEGORYANY_qr(1) = sel_TMTESCODE_rs("TYPE")
                                     'UPD_TMCATEGORY_qr(2) = GETCASE_RS("PATNSTAT")
                                     UPD_TMCATEGORYANY_qr.Execute
                        
                                     FEC_ASSIGN

                                     assign

                                 Else '8
                               
                                     If Len(Trim(Trim(sel_TMCATEGORYANY_rs("SERIES")))) = 1 Then
                                         tempno = "000" & Trim(Trim(sel_TMCATEGORYANY_rs("SERIES")))
                                     ElseIf Len(Trim(Trim(sel_TMCATEGORYANY_rs("SERIES")))) = 2 Then
                                         tempno = "00" & Trim(Trim(sel_TMCATEGORYANY_rs("SERIES")))
                                     ElseIf Len(Trim(Trim(sel_TMCATEGORYANY_rs("SERIES")))) = 3 Then
                                         tempno = "0" & Trim(Trim(sel_TMCATEGORYANY_rs("SERIES")))
                                     End If
                
                                     MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = ST & tempno
                    
                                     'UPDATE PATHNO IN DREQUISITN
                                     upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                     upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                     upddrequisitn.Execute
                
                                     UPD_TMCATEGORYANY_qr(0) = ST & tempno
                                     
                                     If Trim(patstat) = "PR" Or (Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 1) = "R" And Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 2, 1) <> "A") Then
                                        UPD_TMCATEGORYANY_qr(1) = "X"
                                     Else
                                        UPD_TMCATEGORYANY_qr(1) = sel_TMTESCODE_rs("TYPE")
                                        
                                     End If
                                     
                                     UPD_TMCATEGORYANY_qr.Execute
                     
                                     FEC_ASSIGN

                                     assign
                            
                                End If '8
                            Else
                            '****** 08/11/2011 *****
                            
                            sel_TMCATEGORY_qr(0) = Trim(sel_TMTESCODE_rs("TYPE"))
                            sel_TMCATEGORY_qr(1) = "ANY" 'patstat
                            Set sel_TMCATEGORY_rs = sel_TMCATEGORY_qr.OpenResultset(1, 3)
                        
                            If Not sel_TMCATEGORY_rs.EOF Then '3
                                
                                ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)
                    
                                If Left(Trim(sel_TMCATEGORY_rs("PATHNO")), 6) = Trim(ST) Then '4
                                                                             
                                    tmpathno = sel_TMCATEGORY_rs("PATHNO") + 1
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = tmpathno
                                    'Format of pathno is like YYMMDD
                                    upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                    upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                    upddrequisitn.Execute
                                    
                                    'str = "UPDATE DISLIB.TMCATEGORY SET PATHNO=? WHERE TYPE=? AND STATUS=?"
                                    UPD_TMCATEGORY_qr(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                    UPD_TMCATEGORY_qr(1) = sel_TMTESCODE_rs("TYPE")
                                    UPD_TMCATEGORY_qr(2) = "ANY" 'patstat
                                    UPD_TMCATEGORY_qr.Execute
                                    
                                    FEC_ASSIGN
        
                                    assign
                                    
                                Else '4
                                                       
                                    If Len(Trim(Trim(sel_TMCATEGORY_rs("SERIES")))) = 1 Then
                                        tempno = "000" & Trim(Trim(sel_TMCATEGORY_rs("SERIES")))
                                    ElseIf Len(Trim(Trim(sel_TMCATEGORY_rs("SERIES")))) = 2 Then
                                        tempno = "00" & Trim(Trim(sel_TMCATEGORY_rs("SERIES")))
                                    ElseIf Len(Trim(Trim(sel_TMCATEGORY_rs("SERIES")))) = 3 Then
                                        tempno = "0" & Trim(Trim(sel_TMCATEGORY_rs("SERIES")))
                                    Else
                                        tempno = Trim(Trim(sel_TMCATEGORY_rs("SERIES")))
                                    End If
                                
                                'VIDYA
                                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = ST & tempno '+ 1
                                
                                    upddrequisitn(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                                    upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                    upddrequisitn.Execute
                                    
                                    UPD_TMCATEGORY_qr(0) = ST & tempno
                                    UPD_TMCATEGORY_qr(1) = sel_TMTESCODE_rs("TYPE")
                                    UPD_TMCATEGORY_qr(2) = "ANY" 'patstat
                                    UPD_TMCATEGORY_qr.Execute
                                        
                                    FEC_ASSIGN
            
                                    assign
                             
                                End If '4
                            
                            End If '3
                            
                            End If '08/11/2011
                            
                        End If '2
                  End If '1


            End If ' **TM
            
            '***** Shraddha 25/11/2011 to show Sample description *****
            MSFlexGrid1.Cols = 11
            MSFlexGrid1.ColWidth(8) = 5000
             ROW11 = MSFlexGrid1.row
            str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
            Set SEL_DMUL_QR = rdb.CreateQuery("", str)

            SEL_DMUL_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
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
                'TESTDESC = TESTDESC & "| " & vbCrLf & Trim(SEL_MA101_RS("SER_DET"))
                'following line is added by Shweta on 07-05-2012
                ToEnterUnit Trim(SEL_DMUL_RS("PRO_CODE")), ROW11
                SEL_DMUL_RS.MoveNext
            Wend
    
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 8) = TESTDESC
            
            '***********************************************************
   
                
'                '***** Shraddha 20/10/09 *****
'                If Len(txt_pathno.Text) < 3 Then
'                    MsgBox "Please enter 3 digit Lab No"
'                    Exit Sub
'                End If
'                '*****************************


Else '$
            txt_pathno.Visible = False
End If '$
'        End If '*
        '***********************************************
              
    End With
    If DEPTB = "GZZ" Then
    
     str = "SELECT USR_ID FROM DISLIB.DISLNKUSR "
    str = str + "Where USR_ID = '" & Trim(UID) & "' AND (DEPT_CODE='GZZ' OR DEPT_CODE='ALL' OR USRTYPE=9)"
    
    Set REPWTRN_QR = rdb.CreateQuery("", str)
    Set REPWTRNRS = REPWTRN_QR.OpenResultset(1, 1)
    
    If REPWTRNRS.EOF = True Then
    Exit Sub
        End If
    
    
     If cmb_location.Text = "GZZ" Then
          'MSFlexGrid1.row = MSFlexGrid1.row
          'MSFlexGrid1.col = 1
     MSFlexGrid1.ColWidth(1) = 10
     MSFlexGrid1.FormatString = "S.No |^ R|^Case No.  |^ Name            |^M|^Rec Date|^RecTime    |^       "
     End If
    
        If MODALITYB = "SE" Then
        SEAUTONO
        
        Else
        MICROPATH
        End If
        
        
        
        
        
    End If
    
    rdb.BeginTrans
             
    If DEPTB <> "FEC" And (DEPTB <> "FED" And MODALITY <> "BC") Then 'If condition added by Shraddha on 24/12/09
        assign
    End If
    
    
    rdb.CommitTrans

End If
            
    DoEvents
    TXTREQNO.Text = ""
   ' TXTREQNO.Visible = True
    
   ' TXTREQNO.SetFocus 'SHWETA 09-08-2012
    rdb.CommitTrans
    
    
    '***** Check whether data exist in TEMPCOLLCT 13/10/2011 *****
Dim name As String
Dim SEX As String
Dim age As String

 seldmulproce_qr(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 1)
    
While Not seldmulproce_rs.EOF
    
SEL_TEMPCOLLCTR_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
SEL_TEMPCOLLCTR_QR(1) = Trim(seldmulproce_rs("PRO_CODE"))
Set SEL_TEMPCOLLCTR_rs = SEL_TEMPCOLLCTR_QR.OpenResultset(1, 3)
        
If SEL_TEMPCOLLCTR_rs.EOF Then

    If Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))) = 6 Then
            str = "SELECT * FROM Dislib.DREQSTAFFD WHERE REQNO=?" 'DREQSTAFFD
            Set SEL_TA101_QR = rdb.CreateQuery("", str)
            SEL_TA101_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
        Else
            str = "select * from PATOOLSLIB.ta101 where case_no=?"
            Set SEL_TA101_QR = rdb.CreateQuery("", str)
            SEL_TA101_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))
        End If
        Set SEL_TA101_RS = SEL_TA101_QR.OpenResultset(2, 3)
        '*********************
                            
        If Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))) = 6 Then
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
        
        GETCASE_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
        Set GETCASE_RS = GETCASE_QR.OpenResultset(1, 3)
        
        If Not GETCASE_RS.EOF Then
            patnstat = GETCASE_RS("patnstat")
        End If
'INSERT INTO DISLIB.TEMPCOLLCT (SRNO,REQNO,MODALITYCD,CASENO,NAME,SEX,AGE,RECDIPAD,PATNSTAT,PAT_CATG,
'LOCATION,COLLDATE,COLLTIME,SENTDATE,SENTTIME,SER_CODE,SER_DET,SENTFLAG,CFLAG,USERID)

        INS_TEMPCOLLCT(0) = 0 'SRNO
        INS_TEMPCOLLCT(1) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)) 'REQNO
        INS_TEMPCOLLCT(2) = Mid(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)), 4, 2) 'MODALITYCD
        INS_TEMPCOLLCT(3) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)) 'CASENO
        INS_TEMPCOLLCT(4) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3)) 'NAME
        INS_TEMPCOLLCT(5) = SEX
        INS_TEMPCOLLCT(6) = age
        INS_TEMPCOLLCT(7) = ipadd
        INS_TEMPCOLLCT(8) = patnstat
        INS_TEMPCOLLCT(9) = pat_catg
        INS_TEMPCOLLCT(10) = "" 'LOCATION
        INS_TEMPCOLLCT(11) = DATE11 'COLLDATE
        INS_TEMPCOLLCT(12) = "00:00:00" 'COLLTIME
        INS_TEMPCOLLCT(13) = DATE11 'SENTDATE
        INS_TEMPCOLLCT(14) = "00:00:00" 'SENTTIME
               
        'SER_CODE,SER_DET,SENTFLAG,CFLAG
        'seldmulproce_qr(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
        'Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 1)
        
        
        INS_TEMPCOLLCT(15) = Trim(seldmulproce_rs("PRO_CODE")) 'SER_CODE
        
        SELMA101_QR(0) = Trim(seldmulproce_rs("PRO_CODE"))
        Set SELMA101_RS = SELMA101_QR.OpenResultset(1, 3)
        
        If Not SELMA101_RS.EOF Then
            INS_TEMPCOLLCT(16) = Trim(SELMA101_RS("SER_DET"))
        End If
        INS_TEMPCOLLCT(17) = "" 'SENTFLAG
        INS_TEMPCOLLCT(18) = "Y" 'Trim(MSFlexGrid1.TextMatrix(i, 7)) 'CFLAG
        INS_TEMPCOLLCT(19) = Right(Trim(UID), 6) 'USERID
        INS_TEMPCOLLCT(20) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)) 'FLD1 ie pathno
        INS_TEMPCOLLCT.Execute
       
End If
       
        seldmulproce_rs.MoveNext
        Wend
'*******************************************************
If MSFlexGrid1.row > 0 And DEPTB = "FEC" Then
Dim qr_rlocat As rdoQuery
Dim rs_rlocat As rdoResultset

Dim qr_dreq As rdoQuery
Dim rs_dreq As rdoResultset
Set qr_dreq = rdb.CreateQuery("", "select * from dislib.drequisitn where reqno='" & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)) & "'")
Set rs_dreq = qr_dreq.OpenResultset(1, 3)
If rs_dreq.EOF = False Then
Dim WD As String, URG As String
If Trim(rs_dreq("WARD_NAME")) <> "" Or IsNull(Trim(rs_dreq("WARD_NAME"))) = False Then
WD = Trim(rs_dreq("WARD_NAME"))
    Set qr_rlocat = rdb.CreateQuery("", "select * from rislib.rlocation where LOCATION='" & Trim(UCase(WD)) & "' or fld1='" & Trim(UCase(WD)) & "'")
    
    Set rs_rlocat = qr_rlocat.OpenResultset(1, 3)
    If rs_rlocat.EOF = False Then
        If Trim(rs_rlocat("EXTN")) <> "" And IsNull(rs_rlocat("EXTN")) = False Then
            WD = WD & "(" & Trim(rs_rlocat("EXTN")) & ")"
          End If
    End If
Else
WD = ""
End If

If Trim(rs_dreq("SPIHREQNO")) = "U" Then
URG = " /#" & Trim(rs_dreq("SPIHREQNO"))
Else
URG = ""
End If


MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 10) = WD & URG
MSFlexGrid1.row = MSFlexGrid1.row
'MSFlexGrid1.Cols = MSFlexGrid1.row
    If Trim(URG) = "/Urgent" Then
    For KL = 0 To 10
    MSFlexGrid1.col = KL
    MSFlexGrid1.CellBackColor = &H80000013
    Next
    End If
End If
End If
    'MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)
     'If DEPTB = "UZZ" Then 'SHWETA AS PER KALPANA MAM SUGGESTION
        TEMPRECEIVE 'SHWETA 25-09-2012
      '  End If  ''''''
        TXTREQNO.SetFocus
    Exit Sub

MYERR:
    rdb.RollbackTrans 'Shraddha 27/06/10
    MsgBox (Err.Description)

End Sub

Public Sub assign()


'            UPDRREQ_QR(0) = "E"
'            UPDRREQ_QR(1) = cmb_location.Text
'            UPDRREQ_QR(2) = COLLDATE
'            UPDRREQ_QR(3) = REQNOB
'            UPDRREQ_QR.Execute

'update rislib/rrequisitn set examdt,reqnstat,examinatn,location,apntdt,whr reqno
'***** Shraddha 22/09/2011 *****
If DEPTB = "NZZ" Then
    'SELDTMNO_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))
    'Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
    
    SEL_PathDreq_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
    Set SEL_PathDreq_RS = SEL_PathDreq_QR.OpenResultset(1, 3)
    
    If Not SEL_PathDreq_RS.EOF Then
        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = Trim(SEL_PathDreq_RS("PATHNO"))
    End If
End If
'********************************
    updrrequisitn_qr(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    If SELREQNORSFLAG = False Then
        SELREQNOQR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
        Set SELREQNORS = SELREQNOQR.OpenResultset(1, 3)
        SELREQNORSFLAG = True
    Else
        SELREQNOQR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
        SELREQNORS.MoreResults
        SELREQNORS.Requery
    End If
    'FOLLOWING LINE ADDED BY SHWETA 27-03-2012 BCOZ ERR FOR INVALID PARAMETER FOR DEFAULT PARAMETER
    updrrequisitn_qr(1) = Trim(SELREQNORS("reqnstat"))
    '***** Shraddha 09/12/09 *****
    If MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) <> "" Then
        If Trim(SELREQNORS("reqnstat")) = "P" Then
            updrrequisitn_qr(1) = "E"
        End If
        If Trim(SELREQNORS("reqnstat")) = "C" Then
            updrrequisitn_qr(1) = "E"
        End If
            
        If Trim(SELREQNORS("reqnstat")) = "C" Then
            updrrequisitn_qr(1) = "E"
        End If
        
    End If '09/12/09
    
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
    If Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 2) = "RF" Or Left(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)), 2) = "RG" Then
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
    updrrequisitn_qr(5) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
    updrrequisitn_qr.Execute
  '********** End of Update rislib/rrequisitn **************


'* Update rislib/rtimelog set samrecdt,samrectm,caseno,reqndt,reqntm,whr reqno *
    
    
    UPDTIME_QR(0) = Format(txt_date.Text, "YYYY-MM-DD")
    UPDTIME_QR(1) = Format(txt_time.Text, "HH:MM:SS")
    UPDTIME_QR(2) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)) 'case no
    UPDTIME_QR(3) = SELREQNORS("REQNDT")
    UPDTIME_QR(4) = SELREQNORS("REQNTM")
    UPDTIME_QR(5) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1) 'reqno
    UPDTIME_QR.Execute
'************************** End of rislib/rtimelog ***********************
                    
'            UPDRTIME_QR(0) = COLLDATE
'            UPDRTIME_QR(1) = COLLTIME
'            UPDRTIME_QR(2) = REQNOB
'            UPDRTIME_QR.Execute
'***************************************
'       If MODALITY = "BC" Or MODALITY = "EL" Or MODALITY = "TM" Then
'                        If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) > 3 Then
'                            upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
'                        End If
'                        If Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(i, 6))) <> 0 Then
'                            upddrequisitn(0) = ST & UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
'                        End If
'                    Else
'                        upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(i, 6)))
'                    End If
'*******************************************

'27/11/08 only if microbio GZZ
Dim MODALITY As String
Dim hlreqno As String
SELDEPT_QR(0) = MODALITYB
Set SELDEPT_RS = SELDEPT_QR.OpenResultset(1, 3)
MODALITY = Trim(SELDEPT_RS("MODALITYCD"))

'****************
    ' pATHNO
If MODALITY = "SE" Then
    upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))
    upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
    upddrequisitn.Execute
End If

'*******************
'iNTERFACE
If MODALITY = "SE" Then
    hlreqno = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
    seldmulproce_qr(0) = hlreqno
    Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 1)

    Do While Not seldmulproce_rs.EOF
        If SELDBCPROSUBFLAG = False Then
            SELDBCPROSUB_QR1(0) = seldmulproce_rs("PRO_CODE")
            Set SELDBCPROSUB_RS1 = SELDBCPROSUB_QR1.OpenResultset(1, 3)
            SELDBCPROSUBFLAG = True
        Else
            SELDBCPROSUB_QR1(0) = seldmulproce_rs("PRO_CODE")
            SELDBCPROSUB_RS1.MoreResults
            SELDBCPROSUB_RS1.Requery
        End If
        
        If SELDBCPROSUB_RS1.RowCount > 0 Then
            While Not SELDBCPROSUB_RS1.EOF
                SELSAMPLECODE_QR(0) = SELDBCPROSUB_RS1("PROCODESUB")
                Set rs_samplecd = SELSAMPLECODE_QR.OpenResultset(1, 1)
                I = MSFlexGrid1.row
                ins_minterface(0) = MSFlexGrid1.TextMatrix(I, 2)
                ins_minterface(1) = hlreqno
                ins_minterface(2) = Trim(Mid(hlreqno, 4, 10)) & rs_samplecd(0)
                ins_minterface(3) = SELDBCPROSUB_RS1("PROCODESUB")
                ins_minterface(4) = "1"
                ins_minterface(5) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)
                ins_minterface.Execute
                SELDBCPROSUB_RS1.MoveNext
            Wend
        Else
            SELSAMPLECODE_QR(0) = seldmulproce_rs("pro_code")
            Set rs_samplecd = SELSAMPLECODE_QR.OpenResultset(1, 1)
            ins_minterface(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
            ins_minterface(1) = hlreqno
            ins_minterface(2) = Trim(Mid(hlreqno, 4, 10)) & rs_samplecd(0)
            ins_minterface(3) = seldmulproce_rs("pro_code")
            ins_minterface(4) = "1"
            ins_minterface(5) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)
            ins_minterface.Execute
        End If
        seldmulproce_rs.MoveNext
    Loop

'    Do While seldmulproce_rs.EOF = False
'        SELSAMPLECODE_QR(0) = seldmulproce_rs("pro_code")
'        Set rs_samplecd = SELSAMPLECODE_QR.OpenResultset(1, 1)
'        ins_minterface(0) = MSFlexGrid1.TextMatrix(i, 2)
'        ins_minterface(1) = hlreqno
'        ins_minterface(2) = Trim(Mid(hlreqno, 4, 10)) & rs_samplecd(0)
'        ins_minterface(3) = seldmulproce_rs("pro_code")
'        ins_minterface(4) = "1"
'        ins_minterface(5) = MSFlexGrid1.TextMatrix(i, 6)
'        ins_minterface.Execute
'        seldmulproce_rs.MoveNext
'    Loop

End If
       
'***** Shraddha 14/10/09 *****

' HEMATO INTERFACE CODE BEGINS
If MODALITY = "HL" Then
    hlreqno = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
    
    
 seldmulproce_qr(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 1)
    
    If seldmulproce_rs.EOF = False Then
        'FOLLOWING CONDITION IS ADDED BY SHWETA AS TOLD BY KALPANA MAM 25-09-2012
        If seldmulproce_rs("PRO_CODE") = "U722" Or seldmulproce_rs("PRO_CODE") = "U724" Or seldmulproce_rs("PRO_CODE") = "F722" Or seldmulproce_rs("PRO_CODE") = "F724" Then
        
            INSHINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
            INSHINTERFACE_QR(1) = hlreqno
            INSHINTERFACE_QR(2) = " "
            INSHINTERFACE_QR(3) = "1"
                If seldmulproce_rs("PRO_CODE") = "F722" Or seldmulproce_rs("PRO_CODE") = "F724" Then
                INSHINTERFACE_QR(4) = Trim(Mid(hlreqno, 4, 10)) + "E"
                ElseIf seldmulproce_rs("PRO_CODE") = "U722" Then
                 INSHINTERFACE_QR(4) = Trim(Mid(hlreqno, 4, 10)) + "C"
                
                Else
            INSHINTERFACE_QR(4) = Trim(Mid(hlreqno, 4, 10)) + "A"
                End If
            
            INSHINTERFACE_QR(5) = Date
            INSHINTERFACE_QR.Execute
            
            '***** Shraddha 18/11/10 *****
            'INS_IHINTERFACE_QR
            INS_IHINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
            INS_IHINTERFACE_QR(1) = hlreqno
            INS_IHINTERFACE_QR(2) = " "
            INS_IHINTERFACE_QR(3) = "1"
            If seldmulproce_rs("PRO_CODE") = "F722" Or seldmulproce_rs("PRO_CODE") = "F724" Then
            INSHINTERFACE_QR(4) = Trim(Mid(hlreqno, 4, 10)) + "E"
            Else
            INSHINTERFACE_QR(4) = Trim(Mid(hlreqno, 4, 10)) + "A"
            End If
            
            INS_IHINTERFACE_QR(4) = Trim(Mid(hlreqno, 4, 10)) + "A"
            INS_IHINTERFACE_QR(5) = Date
            INS_IHINTERFACE_QR.Execute
        End If
    End If
    
    SELDHEMCYTO_QR(0) = hlreqno
    Set SELDHEMCYTO_RS = SELDHEMCYTO_QR.OpenResultset(1, 1)
    If SELDHEMCYTO_RS.RowCount > 0 Then
   ' Call HINTERFACEDATA(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2), hlreqno)
    
    'FOLLOWING CONDITION COMMENTED BY SHWETA 02-FEB-2012 AS TOLD BY KALPANA MAM
        'If Trim(SELDHEMCYTO_RS("CBC")) = "YES" Or Trim(SELDHEMCYTO_RS("HEM")) = "YES" Then
        'CASENO,REQNO,URGENT,IFLAG,LABNO,TDATE
 'FOLLOWING INSERT COMMENTED BY SHWETA 25-09-2012 AS TOLD BY KALPANA MAM
''''''            INSHINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
''''''            INSHINTERFACE_QR(1) = hlreqno
''''''            INSHINTERFACE_QR(2) = " "
''''''            INSHINTERFACE_QR(3) = "1"
''''''            INSHINTERFACE_QR(4) = Trim(Mid(hlreqno, 4, 10)) + "A"
''''''            INSHINTERFACE_QR(5) = Date
''''''            INSHINTERFACE_QR.Execute
''''''
''''''            '***** Shraddha 18/11/10 *****
''''''            'INS_IHINTERFACE_QR
''''''            INS_IHINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
''''''            INS_IHINTERFACE_QR(1) = hlreqno
''''''            INS_IHINTERFACE_QR(2) = " "
''''''            INS_IHINTERFACE_QR(3) = "1"
''''''            INS_IHINTERFACE_QR(4) = Trim(Mid(hlreqno, 4, 10)) + "A"
''''''            INS_IHINTERFACE_QR(5) = Date
''''''            INS_IHINTERFACE_QR.Execute
            '*****************************
            
       ' End If
            
        If Trim(SELDHEMCYTO_RS("CG")) = "YES" Then
            seldmulproce_qr(0) = hlreqno
            Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 1)
            Do While seldmulproce_rs.EOF = False
                seldhlprosub_qr(0) = seldmulproce_rs("pro_code")
                Set seldhlprosub_rs = seldhlprosub_qr.OpenResultset(1, 1)
                
                'CASENO,REQNO,LABNO,TESTCODE,URGENT,IFLAG,TTIME,TDATE
                
                If seldmulproce_rs("pro_code") = "U709" Or seldmulproce_rs("pro_code") = "U703" Or seldmulproce_rs("pro_code") = "F709" Or seldmulproce_rs("pro_code") = "F703" Then
                    INS_CINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                    INS_CINTERFACE_QR(1) = hlreqno
                    INS_CINTERFACE_QR(2) = Trim(Mid(hlreqno, 4, 10)) + Trim(seldhlprosub_rs("samplecd"))
                    INS_CINTERFACE_QR(3) = Left(seldmulproce_rs("pro_code"), 1) & "001" 'For INR
                    INS_CINTERFACE_QR(4) = ""
                    INS_CINTERFACE_QR(5) = "1"
                    INS_CINTERFACE_QR(6) = time1
                    INS_CINTERFACE_QR(7) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                    INS_CINTERFACE_QR.Execute
                End If
                '******************************************
                'Set seldhlprosub_rs = seldhlprosub_qr.OpenResultset(1, 1) 'Commented & added above
                Do While seldhlprosub_rs.EOF = False
                    'INTERFACE.CINTERFACE(CASENO,REQNO,TESTCODE,URGENT,IFLAG,LABNO)
                    'CASENO , reqno, TESTCODE, URGENT, IFLAG, LABNO
                    If seldhlprosub_rs("samplecd") = "B" Or seldhlprosub_rs("samplecd") = "E" Then
                        INSCINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                        INSCINTERFACE_QR(1) = hlreqno
                        INSCINTERFACE_QR(2) = seldhlprosub_rs("procodesub")
                        INSCINTERFACE_QR(3) = " "
                        INSCINTERFACE_QR(4) = "1"
                        INSCINTERFACE_QR(5) = Trim(Mid(hlreqno, 4, 10)) + "B" ' Trim(seldhlprosub_rs("samplecd"))
                        INSCINTERFACE_QR.Execute
                    End If
                    seldhlprosub_rs.MoveNext
                Loop
                seldmulproce_rs.MoveNext
            Loop
        End If
    End If 'dhemctyo rowcount
    
    '***** Shraddha 02/12/2011 to show Sample description *****
            MSFlexGrid1.Cols = 11
            MSFlexGrid1.ColWidth(8) = 500
            MSFlexGrid1.ColWidth(9) = 5000
            
            MSFlexGrid1.TextMatrix(0, 8) = ""
            MSFlexGrid1.TextMatrix(0, 9) = "Test Description"
            '   MSFlexGrid1.TextMatrix(0, 10) = "Ward/Beds"
                 If DEPTB = "UZZ" And MODALITYB = "HL" Then
                MSFlexGrid1.TextMatrix(0, 8) = "Procedure"
                MSFlexGrid1.TextMatrix(0, 10) = "Status"
               ' MSFlexGrid1.TextMatrix(0, 11) = "Ward/Beds"
               MSFlexGrid1.TextMatrix(0, 9) = ""
               MSFlexGrid1.ColWidth(8) = 5000
            MSFlexGrid1.ColWidth(9) = 500
               
               
                End If
                
            str = "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?"
            Set SEL_DMUL_QR = rdb.CreateQuery("", str)

            SEL_DMUL_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
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
                'TESTDESC = TESTDESC & "| " & vbCrLf & Trim(SEL_MA101_RS("SER_DET"))
                SEL_DMUL_RS.MoveNext
            Wend
    
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 8) = TESTDESC
            
            '***** OPPROCEDUR *****
            SEL_PathDreq_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
            Set SEL_PathDreq_RS = SEL_PathDreq_QR.OpenResultset(1, 3)
            
            If Not SEL_PathDreq_RS.EOF Then
                If IsNull(SEL_PathDreq_RS("OPPROCEDUR")) = False Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 9) = Trim(SEL_PathDreq_RS("OPPROCEDUR"))
                End If
            End If
            '***********************************************************
    
            If SELREQNORS("PATNSTAT") = "IP" Then
             MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 10) = Trim(SEL_PathDreq_RS("WARD_NAME")) & "/" & Trim(SEL_PathDreq_RS("bedno"))
            Else
            MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 10) = Trim(SELREQNORS("PATNSTAT"))
            End If
    
End If  'HL
End Sub
Sub HINTERFACEDATA(ByVal CASENO As String, ByVal REQNO As String)
Dim str1 As String
str1 = "SELECT * FROM DISLIB.HINTERFACE WHERE CASENO=? AND REQNO=?"
Dim selhinterface_qr As rdoQuery
Dim selhinterface_rs As rdoResultset
Set selhinterface_qr = rdb.CreateQuery("", str1)
selhinterface_qr(0) = CASENO
selhinterface_qr(1) = REQNO
Set selhinterface_rs = selhinterface_qr.OpenResultset(1, 3)

If selhinterface_rs.RowCount > 0 Then
str1 = "DELETE FROM DISLIB.HINTERFACE WHERE CASENO='" & Trim(CASENO) & "' AND REQNO='" & Trim(REQNO) & "'"
rdb.Execute str1
End If
End Sub

Public Sub FEC_ASSIGN()

'*****************************
'***** Shraddha 16/10/09 *****
'Update necessary records in tables accordingly
'DInterface dmul
'***** Shraddha 15/10/09 *****
Dim MODALITY As String
SELDEPT_QR(0) = MODALITYB
Set SELDEPT_RS = SELDEPT_QR.OpenResultset(1, 3)
MODALITY = Trim(SELDEPT_RS("MODALITYCD"))

'** 24/12/09 Check here if pathno has been inserted into drequisitn **

SEL_PathDreq_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
Set SEL_PathDreq_RS = SEL_PathDreq_QR.OpenResultset(1, 3)

If Not SEL_PathDreq_RS.EOF Then
    If SEL_PathDreq_RS("PATHNO") = "" Then
        MsgBox "Path No not found.. Please collect sample again"
        UPD_RREQNSTAT_QR(0) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1))
        UPD_RREQNSTAT_QR.Execute
        Exit Sub
    End If
End If
'**********************************************************************
'save

If MODALITY = "BC" And dept = "FEC" Then

    seldmulproce_qr(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
    Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 3)
    
    fecreqno = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
    
    If Not seldmulproce_rs.EOF Then
        Do While Not seldmulproce_rs.EOF
            'checking for record in excption codes which are used as manual and not in interface
            'STORING DATA IN INTERFACE/DINTERFACE
            PCODE = seldmulproce_rs("PRO_CODE")
            If PCODE = "F849" Or PCODE = "F854" Or PCODE = "F924" Or PCODE = "F860" Or PCODE = "F875" Or PCODE = "F880" Or PCODE = "F875" Or PCODE = "F865" Or PCODE = "F802" Or PCODE = "F845" Or PCODE = "F894" Then
                SELDBCPROSUB_QR(0) = PCODE
                Set SELDBCPROSUB_RS = SELDBCPROSUB_QR.OpenResultset(1, 3)
'                Do Until SELDBCPROSUB_RS.EOF
'                    interexpcd_qr(0) = SELDBCPROSUB_RS!PROCODESUB
'                    Set interexpxd_rs = interexpcd_qr.OpenResultset(1, 1)
'                    INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
'
'                    Dim REQ1, P_CODE, SCODE
'                    REQ1 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
'
'
'
'                    P_CODE = SELDBCPROSUB_RS!PROCODESUB
'                    SELSAMPLECODE_QR(0) = P_CODE
'                    Set SAMPLECODE_RS = SELSAMPLECODE_QR.OpenResultset(1, 3)
'
'                    SCODE = SAMPLECODE_RS!SAMPLECD
'                    REQ2 = Mid$(REQ1, 4)
'                    SCODE = REQ2 & SCODE
'                    INSDINTERFACE_QR(1) = REQ1
'                    INSDINTERFACE_QR(2) = SELDBCPROSUB_RS!PROCODESUB 'SHWETA 12-OCT-2012
'                    INSDINTERFACE_QR(3) = ""
'                    INSDINTERFACE_QR(4) = "1"
'                    INSDINTERFACE_QR(5) = SCODE
'
''                    INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
'                    INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
'                    INSDINTERFACE_QR.Execute
'                    SELDBCPROSUB_RS.MoveNext
'                 Loop
Do Until SELDBCPROSUB_RS.EOF
                    interexpcd_qr(0) = SELDBCPROSUB_RS!PROCODESUB
                    Set interexpxd_rs = interexpcd_qr.OpenResultset(1, 1)
                    INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                    Dim REQ1, P_CODE, SCODE
                    REQ1 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                                      
                    
                    If PCODE = "F810" Then
                        If (Left(Trim(seldmulproce_rs("collection")), 1) = Trim(SELDBCPROSUB_RS!flag)) Or (Trim(SELDBCPROSUB_RS!flag) = "B") Then
                       
                        P_CODE = SELDBCPROSUB_RS!PROCODESUB
                        Else
                        GoTo aabc
                        End If
                    Else
                    P_CODE = SELDBCPROSUB_RS!PROCODESUB
                    End If
                    
                    SELSAMPLECODE_QR(0) = P_CODE
                    Set SAMPLECODE_RS = SELSAMPLECODE_QR.OpenResultset(1, 3)
                    
                    SCODE = SAMPLECODE_RS!SAMPLECD
                    REQ2 = Mid$(REQ1, 4)
                    SCODE = REQ2 & SCODE
                    INSDINTERFACE_QR(1) = REQ1
                    INSDINTERFACE_QR(2) = SELDBCPROSUB_RS!PROCODESUB 'SHWETA 12-OCT-2012
                    INSDINTERFACE_QR(3) = ""
                    INSDINTERFACE_QR(4) = "1"
                    INSDINTERFACE_QR(5) = SCODE
                    
'                    INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                    INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                    INSDINTERFACE_QR.Execute
aabc:
                    SELDBCPROSUB_RS.MoveNext
                    
                 Loop
            Else
                interexpcd_qr(0) = seldmulproce_rs("PRO_CODE")
                Set interexpxd_rs = interexpcd_qr.OpenResultset(1, 1)
                If interexpxd_rs.RowCount = 0 Then
                    INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                    Dim REQ11, P_CODE1, SCODE1
                    REQ11 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
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
                    'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                    INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                    On Error GoTo BC1
                    INSDINTERFACE_QR.Execute
BC1:
                    'ADDED BY SHWETA NOV2014 FOR UIBC
                    If seldmulproce_rs("PRO_CODE") = "F893" Then
                    INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                    INSDINTERFACE_QR(1) = REQ11
                    INSDINTERFACE_QR(2) = "F027"
                    INSDINTERFACE_QR(3) = ""
                    INSDINTERFACE_QR(4) = "1"
                    INSDINTERFACE_QR(5) = SCODE1
                    'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                    INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                    INSDINTERFACE_QR.Execute
                    End If
                    
                      If seldmulproce_rs("PRO_CODE") = "F895" Then
                    INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                    INSDINTERFACE_QR(1) = REQ11
                    INSDINTERFACE_QR(2) = "F026"
                    INSDINTERFACE_QR(3) = ""
                    INSDINTERFACE_QR(4) = "1"
                    INSDINTERFACE_QR(5) = SCODE1
                    'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                    INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                    INSDINTERFACE_QR.Execute
                    End If
                    
                End If
            End If
            seldmulproce_rs.MoveNext
        Loop
    End If
End If

'''''' KANCHAN & PRATS  8/01/2008

If MODALITY = "EL" Or ((MODALITY = "BC" And dept <> "FEC") Or dept = "FED") Then

    seldmulproce_qr1(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
    Set seldmulproce_rs1 = seldmulproce_qr1.OpenResultset(1, 3)

    REQNO1 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)

    If Not seldmulproce_rs1.EOF Then

        Do While Not seldmulproce_rs1.EOF

            If seldmulproce_rs1("PRO_CODE") = "F860" Or seldmulproce_rs1("PRO_CODE") = "F875" Or seldmulproce_rs1("PRO_CODE") = "F908" Or seldmulproce_rs1("PRO_CODE") = "F903" Or seldmulproce_rs1("PRO_CODE") = "F949" Or seldmulproce_rs1("PRO_CODE") = "F981" Or seldmulproce_rs1("PRO_CODE") = "F009" Or seldmulproce_rs1("PRO_CODE") = "F901" Or seldmulproce_rs1("PRO_CODE") = "F902" Or seldmulproce_rs1("PRO_CODE") = "F995" Or seldmulproce_rs1("PRO_CODE") = "F022" Then
                 
                seldelprosub_qr(0) = Trim(seldmulproce_rs1("PRO_CODE"))
                Set seldelprosub_rs = seldelprosub_qr.OpenResultset(1, 3)

                If seldelprosub_rs.RowCount > 0 Then

                    While Not seldelprosub_rs.EOF

                        Dim REQ111t, P_CODE111t, SCODE111t
                        REQ111t = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                        SAMPLECODE_QR1(0) = Trim(seldelprosub_rs("PROCODESUB"))
                        Set SAMPLECODE_RS1 = SAMPLECODE_QR1.OpenResultset(1, 3)
                        
                        If SAMPLECODE_RS1.RowCount > 0 Then
                            SCODE111t = SAMPLECODE_RS1!SAMPLECD
                        Else
                            SCODE111t = " "
                        End If
    
                        REQ22 = Mid$(REQ111t, 4)
                        SCODE111t = REQ22 & SCODE111t
    
                        insdinterface_qr1(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                        insdinterface_qr1(1) = REQ111t
                        insdinterface_qr1(2) = SCODE111t
                        insdinterface_qr1(3) = seldelprosub_rs("PROCODESUB")
                        insdinterface_qr1(4) = "1"
                        insdinterface_qr1(5) = time1
                        insdinterface_qr1(6) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                        insdinterface_qr1(7) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                        insdinterface_qr1.Execute
    
                        seldelprosub_rs.MoveNext
                    Wend
                End If 'end if >0
            Else

            ''if not groupcode
                Dim REQ1111t, P_CODE1111t, SCODE1111t
                REQ1111t = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                SAMPLECODE_QR1(0) = Trim(seldmulproce_rs1("PRO_CODE"))
                Set SAMPLECODE_RS1 = SAMPLECODE_QR1.OpenResultset(1, 3)
                If SAMPLECODE_RS1.RowCount > 0 Then
                    SCODE1111t = SAMPLECODE_RS1!SAMPLECD
                Else
                    SCODE1111t = " "
                End If
                REQ222 = Mid$(REQ1111t, 4)
                SCODE1111t = REQ222 & SCODE1111t

                insdinterface_qr1(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                insdinterface_qr1(1) = REQ1111t
                insdinterface_qr1(2) = SCODE1111t
                insdinterface_qr1(3) = seldmulproce_rs1("PRO_CODE")
                insdinterface_qr1(4) = "1"
                insdinterface_qr1(5) = time1
                insdinterface_qr1(6) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
                insdinterface_qr1(7) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                insdinterface_qr1.Execute
            End If

            seldmulproce_rs1.MoveNext
        Loop
    End If
End If


'TM INTERFACE 2005

If MODALITY = "TM" Then
    seldmulproce_qr(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
    Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 3)
   
    fecreqno = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)

    If Not seldmulproce_rs.EOF Then
        Do While Not seldmulproce_rs.EOF
            PCODE = seldmulproce_rs("PRO_CODE")
            
            '***** Shraddha 20/12/2011 *****
            ' FOR SAVING BINTERFACE
            If PCODE = "F977" Then
                Set RS_DMULPROCE = rdb.OpenResultset("SELECT * FROM DISLIB.DMULPROCE WHERE REQNO='" & fecreqno & "' AND PRO_CODE='" & PCODE & "'", 1, 1)
                If RS_DMULPROCE.RowCount > 0 Then
                    Dim req As Variant
                    req = Mid$(fecreqno, 4)
                    'CASENO,REQNO,URGENT,IFLAG,LABNO
                    INSBINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                    INSBINTERFACE_QR(1) = fecreqno
                    INSBINTERFACE_QR(2) = ""
                    INSBINTERFACE_QR(3) = "1"
                    INSBINTERFACE_QR(4) = req & "A"
                    SEL_BINTER_QR(0) = fecreqno 'Trim(PCODE)
                    Set SEL_BINTER_RS = SEL_BINTER_QR.OpenResultset(1, 3)
                    If SEL_BINTER_RS.EOF Then
                      
                         INSBINTERFACE_QR.Execute
                    End If
                End If
            End If
            '*******************************
            'SHWETA 30-JAN-2013
'            Dim NEWCFU_TMTESTCODE As Boolean
'            Dim QRNEWCFU_TMTESTCODE As rdoQuery
'            Dim RSNEWCFU_TMTESTCODE As rdoResultset
'            str = "SELECT * FROM DISLIB.TMTESTCODE WHERE TYPE='N' AND SER_CODE = ? "
'            Set QRNEWCFU_TMTESTCODE = rdb.CreateQuery("", str)
'            QRNEWCFU_TMTESTCODE(0) = PCODE
'            Set RSNEWCFU_TMTESTCODE = QRNEWCFU_TMTESTCODE.OpenResultset(1, 3)
'            NEWCFU_TMTESTCODE = False
'            If RSNEWCFU_TMTESTCODE.EOF = False Then
'            NEWCFU_TMTESTCODE = True
'            End If
            
            If PCODE = "F053" Or PCODE = "F054" Or PCODE = "F043" Or PCODE = "F044" Or PCODE = "F045" Or PCODE = "F046" Or PCODE = "F047" Or PCODE = "F0468" Or PCODE = "F049" Or PCODE = "F050" Or PCODE = "F051" Or PCODE = "F052" Then
            'FOR EINTERFACE
            Dim INS_EINTERFACE_QR As rdoQuery
            Dim SEL_EINTERFACE_QR As rdoQuery
            Dim SEL_EINTERFACE_RS As rdoResultset
            str = "INSERT INTO DISLIB.EINTERFACE(REQNO,CASENO,LABNO,PATHNO,CALDATE,CALTIME) VALUES(?,?,?,?,?,?)"
            Set INS_EINTERFACE_QR = rdb.CreateQuery("", str)

            INS_EINTERFACE_QR(0) = fecreqno
            INS_EINTERFACE_QR(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
            INS_EINTERFACE_QR(2) = Trim(UCase(TXTREQNO.Text))
            INS_EINTERFACE_QR(3) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)
            INS_EINTERFACE_QR(4) = Format(DATE11, "YYYY-MM-DD")
            INS_EINTERFACE_QR(5) = time1
            '  INS_EINTERFACE_QR.Execute
              
              str = "SELECT * FROM DISLIB.EINTERFACE WHERE REQNO=?"
              Set SEL_EINTERFACE_QR = rdb.CreateQuery("", str)
              SEL_EINTERFACE_QR(0) = fecreqno 'Trim(PCODE)
                    Set SEL_EINTERFACE_RS = SEL_EINTERFACE_QR.OpenResultset(1, 3)
                    If SEL_EINTERFACE_RS.EOF Then
                      
                         INS_EINTERFACE_QR.Execute
                    End If
              
            
            End If
            
'            If PCODE = "F046" Or PCODE = "F047" Or PCODE = "F0468" Or PCODE = "F049" Or PCODE = "F050" Or PCODE = "F051" Then
'
'
'            End If
            
            '**********************************
            
            If PCODE = "F033" Or PCODE = "F043" Or PCODE = "F045" Or PCODE = "F049" Or PCODE = "F072" Or PCODE = "F076" Or PCODE = "F087" Or PCODE = "F091" Or PCODE = "F109" Or PCODE = "F113" Then
                SELDBCPROSUB_QR(0) = PCODE
                Set SELDBCPROSUB_RS = SELDBCPROSUB_QR.OpenResultset(1, 3)
                Do Until SELDBCPROSUB_RS.EOF
                    interexpcd_qr(0) = SELDBCPROSUB_RS!PROCODESUB
                    Set interexpxd_rs = interexpcd_qr.OpenResultset(1, 1)
                    
                    INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                    Dim REQ1T, P_CODET, SCODET
                    REQ1T = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
                    P_CODET = SELDBCPROSUB_RS!PROCODESUB
                    SELSAMPLECODE_QR(0) = P_CODET
                    Set SAMPLECODE_RS = SELSAMPLECODE_QR.OpenResultset(1, 3)
                    SCODET = SAMPLECODE_RS!SAMPLECD
                    REQ2 = Mid$(REQ1T, 4)
                    SCODET = REQ2 & SCODET
    
                    INSDINTERFACE_QR(1) = REQ1T
                    INSDINTERFACE_QR(2) = SELDBCPROSUB_RS("PROCODESUB")
                    INSDINTERFACE_QR(3) = ""
                    INSDINTERFACE_QR(4) = "1"
                    INSDINTERFACE_QR(5) = SCODET
                    'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                    'Commented and added by Shraddha on 26/08/2011
                    INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                    INSDINTERFACE_QR.Execute
    
                    SELDBCPROSUB_RS.MoveNext
                Loop
            Else
                interexpcd_qr(0) = seldmulproce_rs("PRO_CODE")
                Set interexpxd_rs = interexpcd_qr.OpenResultset(1, 1)
            
                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                Dim REQ11t, P_CODE1t, SCODE1t
                REQ11t = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
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
                'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                'Commented and added by Shraddha on 26/08/2011
                INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                INSDINTERFACE_QR.Execute
                
                
                  If seldmulproce_rs("PRO_CODE") = "F052" Then
                  Dim IFESTR As String
                  IFESTR = "UPDATE INTERFACE.DINTERFACE SET TESTRESULT ='-' WHERE REQNO='" & REQ11t & "' AND TESTCODE='F052'"
                  rdb.Execute IFESTR
                  
                  End If
                
                If seldmulproce_rs("PRO_CODE") = "F057" Or seldmulproce_rs("PRO_CODE") = "F056" Then
                
                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                INSDINTERFACE_QR(1) = REQ11t
                INSDINTERFACE_QR(2) = "F066"
                INSDINTERFACE_QR(3) = ""
                INSDINTERFACE_QR(4) = "1"
                INSDINTERFACE_QR(5) = SCODE1t
                'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                'Commented and added by Shraddha on 26/08/2011
                INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                INSDINTERFACE_QR.Execute
                
                
                 
                End If
                
                
                If seldmulproce_rs("PRO_CODE") = "F064" Or seldmulproce_rs("PRO_CODE") = "F065" Then
                
                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                INSDINTERFACE_QR(1) = REQ11t
                INSDINTERFACE_QR(2) = "F069"
                INSDINTERFACE_QR(3) = ""
                INSDINTERFACE_QR(4) = "1"
                INSDINTERFACE_QR(5) = SCODE1t
                'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                'Commented and added by Shraddha on 26/08/2011
                INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                INSDINTERFACE_QR.Execute
                
                
                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                INSDINTERFACE_QR(1) = REQ11t
                INSDINTERFACE_QR(2) = "F068"
                INSDINTERFACE_QR(3) = ""
                INSDINTERFACE_QR(4) = "1"
                INSDINTERFACE_QR(5) = SCODE1t
                'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                'Commented and added by Shraddha on 26/08/2011
                INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                INSDINTERFACE_QR.Execute
                
                End If
                
                If PCODE = "F050" Then
                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                INSDINTERFACE_QR(1) = REQ11t
                INSDINTERFACE_QR(2) = "F066"
                INSDINTERFACE_QR(3) = ""
                INSDINTERFACE_QR(4) = "1"
                INSDINTERFACE_QR(5) = SCODE1t
                'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                'Commented and added by Shraddha on 26/08/2011
                INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                INSDINTERFACE_QR.Execute
                
                End If
                If PCODE = "F041" Then
                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                INSDINTERFACE_QR(1) = REQ11t
                INSDINTERFACE_QR(2) = "F067"
                INSDINTERFACE_QR(3) = ""
                INSDINTERFACE_QR(4) = "1"
                INSDINTERFACE_QR(5) = SCODE1t
                'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                'Commented and added by Shraddha on 26/08/2011
                INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                INSDINTERFACE_QR.Execute
                   End If
                
                
                If seldmulproce_rs("PRO_CODE") = "F053" Then
                
                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                INSDINTERFACE_QR(1) = REQ11t
                INSDINTERFACE_QR(2) = "F070"
                INSDINTERFACE_QR(3) = ""
                INSDINTERFACE_QR(4) = "1"
                INSDINTERFACE_QR(5) = SCODE1t
                'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                'Commented and added by Shraddha on 26/08/2011
                INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                INSDINTERFACE_QR.Execute
                
                
                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                INSDINTERFACE_QR(1) = REQ11t
                INSDINTERFACE_QR(2) = "F071"
                INSDINTERFACE_QR(3) = ""
                INSDINTERFACE_QR(4) = "1"
                INSDINTERFACE_QR(5) = SCODE1t
                'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                'Commented and added by Shraddha on 26/08/2011
                INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                INSDINTERFACE_QR.Execute
                
                End If
                
                If seldmulproce_rs("PRO_CODE") = "F977" Then
                End If
                
                
                If seldmulproce_rs("PRO_CODE") = "F977" Then
                
                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                INSDINTERFACE_QR(1) = REQ11t
                INSDINTERFACE_QR(2) = "F064"
                INSDINTERFACE_QR(3) = ""
                INSDINTERFACE_QR(4) = "1"
                INSDINTERFACE_QR(5) = SCODE1t
                'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                'Commented and added by Shraddha on 26/08/2011
                INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                INSDINTERFACE_QR.Execute
                
                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                INSDINTERFACE_QR(1) = REQ11t
                INSDINTERFACE_QR(2) = "F065"
                INSDINTERFACE_QR(3) = ""
                INSDINTERFACE_QR(4) = "1"
                INSDINTERFACE_QR(5) = SCODE1t
                'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                'Commented and added by Shraddha on 26/08/2011
                INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                INSDINTERFACE_QR.Execute
                
                'F966 F967 F968 F969 F970 F971
                'CHECK HERE
                Dim ARR_PR(4) As String
                'ARR_PR(0) = "F966"
                'ARR_PR(1) = "F967"
                ARR_PR(0) = "F068"
                ARR_PR(1) = "F069"
                ARR_PR(2) = "F070"
                ARR_PR(3) = "F071"
                
                Dim I_PR As Integer
                For I_PR = 0 To UBound(ARR_PR) - 1
                INSDINTERFACE_QR(0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2)
                INSDINTERFACE_QR(1) = REQ11t
                INSDINTERFACE_QR(2) = ARR_PR(I_PR)
                INSDINTERFACE_QR(3) = ""
                INSDINTERFACE_QR(4) = "1"
                INSDINTERFACE_QR(5) = SCODE1t
                'INSDINTERFACE_QR(6) = Format(DATE11, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                'Commented and added by Shraddha on 26/08/2011
                INSDINTERFACE_QR(6) = Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))
                INSDINTERFACE_QR.Execute
                Next
                
                Else
'                INSDINTERFACE_QR.Execute
                End If
                
                
                
                
                
            End If
            seldmulproce_rs.MoveNext
        Loop
    End If
End If

'*********** 16/10/09 ******************
End Sub


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



Private Sub txt_pathno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
    If MSFlexGrid1.row <> 1 Then
    MSFlexGrid1.row = MSFlexGrid1.row - 1
    txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
    End If
End If
If KeyCode = 40 Then
    If MSFlexGrid1.row <> MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.row = MSFlexGrid1.row + 1
        txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
        txt_pathno = CStr(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
    End If
End If

If KeyCode = 13 Then 'If enter KEY has been pressed

    '***** Shraddha 20/10/09 *****
    'If Len(txt_pathno.Text) < 3 Then
    If Len(txt_pathno.Text) < 4 Then
        'MsgBox "Please enter 3 digit Lab No"
        MsgBox "Please enter 4 digit Lab No"
        Exit Sub
    End If
    '*****************************

    ST = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2)

    'DISPLAY Path No
    
    '***** If PathNo already given 16/10/09 *****
    '*****
    Dim MODALITY As String
    Dim hlreqno As String
    SELDEPT_QR(0) = MODALITYB
    Set SELDEPT_RS = SELDEPT_QR.OpenResultset(1, 3)
    MODALITY = Trim(SELDEPT_RS("MODALITYCD"))
    
    If Trim(txt_pathno.Text) <> "" Then
        If MODALITY = "BC" Or MODALITY = "EL" Or MODALITY = "TM" Then
            For I = 1 To MSFlexGrid1.Rows - 1
                If Trim(txt_pathno.Text) = Trim(MSFlexGrid1.TextMatrix(I, 6)) Then
                    MsgBox "Lab No. Already exists"
                    Exit Sub
                End If
            Next
        
            If Len(Trim(txt_pathno.Text)) <= 4 And Len(Trim(txt_pathno.Text)) <> 0 Then
                SELDREQUISITNQR(0) = ST & UCase(Trim(txt_pathno.Text))
            End If
            If Len(Trim(txt_pathno.Text)) > 4 Then
                SELDREQUISITNQR(0) = UCase(Trim(txt_pathno.Text))
            End If
            SELDREQUISITNQR(1) = Trim(dept) & MODALITY & "%"
            Set SELDREQUISITNRS = SELDREQUISITNQR.OpenResultset(1, 3)
        End If

            If SELDREQUISITNRS.RowCount > 0 Then
                MsgBox "The Path No. is  Given already " & SELDREQUISITNRS("caseno") & vbCrLf & "Sample not yet received, Please give different Path No"
                'txt_pathno.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                txt_pathno.Visible = True '11/12/09
                txt_pathno.Top = 3480
                txt_pathno.Left = 12720
                txt_pathno.Text = ""
                Exit Sub
            End If
    End If

'*****
    '********************************************

    If MODALITYB = "BC" Or MODALITYB = "EL" Or MODALITYB = "TM" Then
        '    If Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))) > 3 Then
        '        upddrequisitn(0) = UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))
        '    End If
        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = txt_pathno.Text
        'If Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))) <> 0 Then
        '***** Commented and added by Shraddha on 29/01/10 *****
        If Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))) <= 4 And Len(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7))) <> 0 Then
            upddrequisitn(0) = ST & UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))
            upd_TEMMPCOLL_QR(0) = ST & UCase(Trim(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)))
        End If
        upddrequisitn(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
        upd_TEMMPCOLL_QR(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1)
        upddrequisitn.Execute
        upd_TEMMPCOLL_QR.Execute
        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = txt_pathno.Text
        txt_pathno.Visible = False

        'Increment and update for next records
        'PathNo=pathno+1 7th col
    
        pathno = MSFlexGrid1.TextMatrix(1, 7)
        If MSFlexGrid1.row > 1 Then
            pathno = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)
            pathno = pathno + 1
        End If
        
        FEC_ASSIGN
        
        
        assign 'Added by Shraddha on 24/12/09
        
        TXTREQNO.SetFocus 'Added on 20/10/09
        
    End If
End If

End Sub

Private Sub TXT_PATHNO_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48) And (KeyAscii <= 57) Or (KeyAscii = 8) Then
    Else
    KeyAscii = 0
    End If
End Sub


Private Sub TxtUnit_KeyPress(KeyAscii As Integer)
If ((KeyAscii >= 48) And (KeyAscii <= 57)) Or (KeyAscii = 8) Or (KeyAscii = 32) Or ((KeyAscii >= 97) And (KeyAscii <= 122)) Or ((KeyAscii >= 65) And (KeyAscii <= 90)) Then
Else
    
If KeyAscii = 13 Then
 MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 9) = TxtUnit.Text
 'MsgBox "Entered Values is : " & TxtUnit.Text
  updatedreq12
 TxtUnit.Visible = False
 TxtUnit.Text = ""
End If
   
    KeyAscii = 0
    End If

End Sub


Sub updatedreq12()
Dim upddrequisitn1 As rdoQuery
Dim str1 As String
str1 = "UPDATE DISLIB.DREQUISITN SET OPPROCEDUR=? WHERE REQNO=?"
Set upddrequisitn1 = rdb.CreateQuery("", str1)
upddrequisitn1(0) = TxtUnit.Text
upddrequisitn1(1) = Replace(REQNOB, "/", "")
upddrequisitn1.Execute
End Sub

