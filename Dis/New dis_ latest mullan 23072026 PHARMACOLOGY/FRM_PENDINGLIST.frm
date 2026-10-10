VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FRM_PENDINGLIST 
   Caption         =   "Form1"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   12315
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR1 
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   10470
      TabIndex        =   27
      Top             =   720
      Width           =   1215
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
      Left            =   10470
      TabIndex        =   26
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   0
      TabIndex        =   22
      Top             =   7080
      Width           =   11775
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
         MouseIcon       =   "FRM_PENDINGLIST.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   360
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
         MouseIcon       =   "FRM_PENDINGLIST.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   360
         Width           =   1095
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
         MouseIcon       =   "FRM_PENDINGLIST.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
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
         MouseIcon       =   "FRM_PENDINGLIST.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3495
      Left            =   0
      TabIndex        =   20
      Top             =   3480
      Width           =   11775
      Begin MSFlexGridLib.MSFlexGrid msf_reqstat 
         Height          =   3255
         Left            =   0
         TabIndex        =   21
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
         FormatString    =   $"FRM_PENDINGLIST.frx":0C28
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   11775
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
         TabIndex        =   14
         Top             =   240
         Width           =   6255
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
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   735
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
         TabIndex        =   2
         Top             =   570
         Width           =   675
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
         TabIndex        =   7
         Top             =   1170
         Width           =   2295
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
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
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
         TabIndex        =   1
         Top             =   600
         Width           =   645
      End
      Begin VB.Frame emppatient 
         Height          =   1095
         Left            =   1080
         TabIndex        =   10
         Top             =   480
         Width           =   1455
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
            TabIndex        =   12
            Top             =   600
            Width           =   975
         End
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
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
      End
      Begin MSMask.MaskEdBox MaskEd_to_date 
         Height          =   375
         Left            =   9120
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   6
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
         TabIndex        =   19
         Top             =   240
         Width           =   630
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
         TabIndex        =   18
         Top             =   600
         Width           =   1140
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
         TabIndex        =   17
         Top             =   1200
         Width           =   1635
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
         TabIndex        =   16
         Top             =   600
         Width           =   1455
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
         TabIndex        =   15
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1635
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   11775
      Begin VB.PictureBox CRYSTALREPORT1 
         Height          =   480
         Left            =   1080
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   31
         Top             =   480
         Width           =   1200
      End
      Begin VB.PictureBox Cr_reqstat 
         Height          =   480
         Left            =   360
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   32
         Top             =   480
         Width           =   1200
      End
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
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   11655
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
      Left            =   9840
      TabIndex        =   30
      Top             =   720
      Width           =   510
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
      Left            =   9840
      TabIndex        =   29
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label Label3 
      Caption         =   "PENDING LIST"
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
      Left            =   4680
      TabIndex        =   28
      Top             =   840
      Width           =   8415
   End
End
Attribute VB_Name = "FRM_PENDINGLIST"
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
                                       Else
                                        If Trim(rs_grid_det("RECPTNO")) = "NOCHARGE" Or Trim(rs_grid_det("RECPTNO")) = "COMPANY" Or Trim(rs_grid_det("RECPTNO")) = "TRUST" Or Trim(rs_grid_det("RECPTNO")) = "IPATIENT" Or Trim(rs_grid_det("RECPTNO")) = "ITRUST" Or Trim(rs_grid_det("RECPTNO")) = "ICOMPANY" Or Trim(rs_grid_det("RECPTNO")) = "VPP" Or Trim(rs_grid_det("RECPTNO")) = "LR" Or Trim(rs_grid_det("RECPTNO")) = "BAR" Or Trim(rs_grid_det("RECPTNO")) = "DrOp" Then
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
                                    If Trim(rs_grid_det(1)) = "T" Then
                                      If IsNull(Trim(rs_grid_det("RECPTNO"))) Then
                                        
                                        appt = "Yes"
                                        Memo = "Yes"
                                        payment = "NO"
                                        EXAM = "Yes"
                                        REPORT = "Yest"
                                      Else
                                      If Trim(rs_grid_det("RECPTNO")) = "NOCHARGE" Or Trim(rs_grid_det("RECPTNO")) = "COMPANY" Or Trim(rs_grid_det("RECPTNO")) = "TRUST" Or Trim(rs_grid_det("RECPTNO")) = "IPATIENT" Or Trim(rs_grid_det("RECPTNO")) = "ITRUST" Or Trim(rs_grid_det("RECPTNO")) = "ICOMPANY" Or Trim(rs_grid_det("RECPTNO")) = "VPP" Or Trim(rs_grid_det("RECPTNO")) = "LR" Or Trim(rs_grid_det("RECPTNO")) = "BAR" Or Trim(rs_grid_det("RECPTNO")) = "DrOp" Then
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
                             msf_reqstat.TextMatrix(srno, 2) = msk_caseno.FormattedText

                             
                             msf_reqstat.col = 3
                             If UCase(Memo) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                            msf_reqstat.TextMatrix(srno, 3) = Trim(rs_grid_det("MEMONO")) & " " & Trim(rs_grid_det("MEMODT")) & ""
                             End If
                            msf_reqstat.col = 4
                             If UCase(payment) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                               msf_reqstat.TextMatrix(srno, 4) = rs_grid_det("RECPTDT") & ""
                             Else
                               msf_reqstat.CellBackColor = vbGreen
                               If Trim(rs_grid_det("Recptno")) = "NOCHARGE" Then
                                msf_reqstat.TextMatrix(srno, 4) = "NOCHARGE"
                               End If

                               If Trim(rs_grid_det("RECPTNO")) = "TRUST" Then
                                msf_reqstat.TextMatrix(srno, 4) = "TRUST"
                               End If
                               
                               If Trim(rs_grid_det("RECPTNO")) = "COMPANY" Then
                                msf_reqstat.TextMatrix(srno, 4) = "COMPANY"
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "IPATIENT" Then
                                msf_reqstat.TextMatrix(srno, 4) = "IPATIENT"
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "ITRUST" Then
                                msf_reqstat.TextMatrix(srno, 4) = "ITRUST"
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "ICOMPNY" Then
                                msf_reqstat.TextMatrix(srno, 4) = "ICOMPANY"
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "VPP" Then
                                msf_reqstat.TextMatrix(srno, 4) = "VPP"
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "LR" Then
                                msf_reqstat.TextMatrix(srno, 4) = "LR"
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "BAR" Then
                                msf_reqstat.TextMatrix(srno, 4) = "BAR"
                               End If
                               If Trim(rs_grid_det("RECPTNO")) = "DrOp" Then
                                msf_reqstat.TextMatrix(srno, 4) = "DrOp"
                               End If
                                
                               If IsNull(Trim(rs_grid_det("RECPTNO"))) Then
                                msf_reqstat.CellBackColor = vbWhite
                                msf_reqstat.TextMatrix(srno, 4) = " "
                               End If
                                                    


                             End If
                             msf_reqstat.col = 5
                             If UCase(appt) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                             msf_reqstat.TextMatrix(srno, 5) = rs_grid_det("APNTDT") & ""
                             End If
                             
                             msf_reqstat.col = 6
                             If UCase(EXAM) = "YES" Then
                             msf_reqstat.CellBackColor = vbGreen
                              If IsNull(Trim(rs_grid_det("EXAMDT"))) Then
                                msf_reqstat.CellBackColor = vbWhite
                                msf_reqstat.TextMatrix(srno, 6) = ""
                              Else
                                    msf_reqstat.TextMatrix(srno, 6) = rs_grid_det("EXAMDT") & ""
                               End If
                             End If
                            msf_reqstat.col = 7
                             If UCase(REPORT) = "YESF" Then
                             msf_reqstat.TextMatrix(srno, 7) = "FNL " & rs_grid_det("FREPORTDT") & ""
                             msf_reqstat.CellBackColor = vbGreen
                             End If
                            
                              If UCase(REPORT) = "YEST" Then
                             msf_reqstat.TextMatrix(srno, 7) = "PRO " & rs_grid_det("PREPORTDT") & ""
                             msf_reqstat.CellBackColor = vbGreen
                             End If
                        
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
                                SELDREQNOQR1(0) = rs_grid_det("REQNNO")
                                Set SELDREQNORS1 = SELDREQNOQR1.OpenResultset(1, 3)
                                If SELDREQNORS1.RowCount > 0 Then
                                    If Len(Trim(SELDREQNORS1("REQNO"))) > 0 Then
                                        msf_reqstat.CellBackColor = vbGreen
                                        msf_reqstat.TextMatrix(srno, 8) = req_chglen(SELDREQNORS1("REQNO"))
                                        
                                        
                                    End If
                                End If
                            End If
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
Dim VAR_REQNNO As String
Dim var_yr_f As Integer, var_yr_t As Integer
Dim var_mon_f As Integer, var_mon_t As Integer
Dim var_day_f As Integer, var_day_t As Integer
Dim var_date_f As String, var_date_t As String



VAR_REQNNO = req_chglen(msf_reqstat.Text)

'CR1.Connect = "UID=PATOOLS;PWD=PATOOLS;Database = EC_TMH"

If Opt_case.Value = True Then
    CR1.SelectionFormula = "{RREQUISITN.CASENO} = '" & Trim(cno_pad) & "'"
    CR1.ReportFileName = "C:\DISREPORTS\reqsts.rpt"
    
     str1 = ip
     str2 = DATE11
     STR3 = time1
    str1 = str1 & Space(5) & str2 & Space(5) & STR3
    str1 = Space(3) & str1
    CR1.Formulas(0) = "us_id = '" & str1 & "'"
Else
    var_date_f = Year(MaskEd_from_date) & "," & month(MaskEd_from_date) & "," & Day(MaskEd_from_date)
    var_date_t = Year(MaskEd_to_date) & "," & month(MaskEd_to_date) & "," & Day(MaskEd_to_date)
''
''   'CR1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date (2001,09 , 25) AND {RREQUISITN.REQNDT} <= Date (2001,12 ,30 )"
''    CR1.SelectionFormula = "{RREQUISITN.MODALITYCD}= '" & Trim(cbo_modality) & "'"
''    MsgBox CR1.SelectionFormula
''    'CR1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date (2001,09 , 25) AND {RREQUISITN.REQNDT} <= Date (2001,12 ,30 )"
''    CR1.Formulas(0) = "curtime = '" & txt_time.Text & "'"
''    CR1.Formulas(1) = "formdate = '" & MaskEd_from_date.FormattedText & "'"
''    CR1.Formulas(2) = "todate = '" & MaskEd_to_date.FormattedText & "'"
''    CR1.ReportFileName = App.Path & "\crreqsts.rpt"
     CR1.ReportFileName = "C:\DISREPORTSEC\reqsts1.rpt"
        str1 = ip
        str2 = DATE11
        STR3 = time1
        str1 = str1 & Space(5) & str2 & Space(5) & STR3
        STR5 = "F"
  'CR3.SelectionFormula = " {TA102.case_no} = '" & MCNO & "' and {TA102.APPT_DATE} = date( " & adt & ") and {ta102.inv_type} = '" & INVTP & "'"   ''modified 23899
        
      
       ' CR1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date (2001,01 , 01) AND {RREQUISITN.REQNDT} <= Date (2002,01 ,01 )"
         CR1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & Trim(cbo_modality) & " ' AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
        'CR1.SelectionFormula = "{RREQUISITN.MODALITYCD}='" & Trim(cbo_modality) & "'"
        'CR1.SelectionFormula = "{RREQUISITN.REQNSTAT} <> '" & "F" & "'"
        
        CR1.Formulas(3) = "us_id = '" & str1 & "'"
        CR1.Formulas(0) = "fromdt = '" & MaskEd_from_date.FormattedText & "'"
        CR1.Formulas(2) = "todt = '" & MaskEd_to_date.FormattedText & "'"
       
End If


CR1.WindowState = crptMaximized
'CR1.Destination =crptToPrinter
CR1.Action = 1
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
  TXT_NAME.Visible = False
  LBL_NAME.Visible = False ' to clear all the text fields on the form
  
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
Opt_all.Enabled = True
Opt_case.Enabled = True
Opt_all.Value = False
Opt_case.Value = False
msk_caseno.Enabled = False
cbo_req.Enabled = False
MaskEd_from_date.Enabled = False
MaskEd_to_date.Enabled = False
End Sub

Private Sub cmd_view1_Click()

End Sub

Private Sub Form_Activate()
'cmd_view.SetFocus
End Sub

Private Sub Form_Load()
Dim str As String
'Connection
TXT_NAME.Visible = False
LBL_NAME.Visible = False
srno = 1
Set RS = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107") '***2/2/99
DATE11 = RS("dt1") '**add 24/11/2k
time1 = Format(RS("TM1"), "hh:mm:ss") '***2/2/99
txt_date = Format(DATE11, "dd/mm/yyyy")
txt_time = time1
'   MaskEd_from_date.Text = Format(DATE11, "dd/mm/yyyy")
'   MaskEd_to_date.Text = Format(DATE11, "dd/mm/yyyy")
exam_flag = False
If exam_flag = False Then
   
    ''DTM BLANK REPORT CHEQUE
'    STR = "SELECT * FROM DISLIB.DTMTAB WHERE REQNO=?"
'    Set SELDTMTAB_QR = rdb.CreateQuery("SELDTMTAB_STR11", STR)
'
    
    
    Frame4.Enabled = True
    'txt_msg.Enabled = False
    cmd_cancel.Enabled = False
    cmd_print.Enabled = False
    txt_date.Enabled = False
    txt_time.Enabled = False
    MaskEd_from_date.Enabled = False
    MaskEd_to_date.Enabled = False
    ' for selecting case no into combo box from RREQUISITN
    Set rq_modcd = rdb.CreateQuery("rq_modcd11", "select distinct(modalitycd) from rislib.rmodality WHERE DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='FZZ'OR DEPT_CODE='FEC'OR DEPT_CODE='LZZ' OR DEPT_CODE='GZZ' OR DEPT_CODE='IZZ'OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' ")
    Set rs_modcd = rq_modcd.OpenResultset(1, 1)
    cbo_modality.clear
    cbo_modality.AddItem "ALL"
    While rs_modcd.EOF = False
        cbo_modality.AddItem rs_modcd(0)
        rs_modcd.MoveNext
    Wend
 
    ' for selecting req no for the selected case no from RREQUISITN
    Set rq_req = rdb.CreateQuery("req1", "select REQNNO,reqndt from Rislib.RREQUISITN where caseno=? AND REQNSTAT='E' order by reqndt desc")
    rq_req(0) = "aaa"
    Set rs_req = rq_req.OpenResultset()
    
    ' for selecting  req no for option all between the dates
    ' Set rq_allreq = rdb.CreateQuery("allreq", "select REQNNO,REQNSTAT,APNTDT,MEMONO,RECPTNO,CASENO from RREQUISITN where REQNDT between cdate(?) and cdate(?)")'????
     Set rq_allreq = rdb.CreateQuery("allreq2", "select REQNNO,REQNSTAT,APNTDT,MEMONO,RECPTNO,CASENO,MEMODT,RECPTDT,EXAMDT,FREPORTDT,PREPORTDT,RECPTNO,reqndt from Rislib.RREQUISITN where reqnstat='E' AND REQNDT between ? and  ? and modalitycd=? order by reqndt desc")
     Set RQ_ALLREQ1 = rdb.CreateQuery("allreq21", "select REQNNO,REQNSTAT,APNTDT,MEMONO,RECPTNO,CASENO,MEMODT,RECPTDT,EXAMDT,FREPORTDT,PREPORTDT,RECPTNO,reqndt from Rislib.RREQUISITN where reqnstat='E' AND REQNDT between ? and  ? order by reqndt desc")
     rq_allreq(0) = "1999/01/01"  '"01/01/1999"
     rq_allreq(1) = "1999/01/10"
     rq_allreq(2) = "CT"
     Set rs_allreq = rq_allreq.OpenResultset(2, 3)
     
     ''FOR SP REPORT
     Set RQ_ALLSP = rdb.CreateQuery("allSP", "select REQNNO,REQNSTAT,APNTDT,MEMONO,RECPTNO,CASENO,MEMODT,RECPTDT,EXAMDT,FREPORTDT,PREPORTDT,RECPTNO,reqndt from Rislib.RREQUISITN where reqnstat='E' AND REQNDT between ? and  ? and modalitycd=? order by reqndt desc")
     Set RQ_ALLSP1 = rdb.CreateQuery("allSP1", "select REQNNO,REQNSTAT,APNTDT,MEMONO,RECPTNO,CASENO,MEMODT,RECPTDT,EXAMDT,FREPORTDT,PREPORTDT,RECPTNO,reqndt from Rislib.RREQUISITN where  reqnstat='E' AND REQNDT between ? and  ? order by reqndt desc")
     RQ_ALLSP(0) = "1999/01/01"  '"01/01/1999"
     RQ_ALLSP(1) = "1999/01/10"
     RQ_ALLSP(2) = "CT"
     Set RS_ALLSP = RQ_ALLSP.OpenResultset(2, 3)
     
     
     
     
    '  for selecting flexgrid details from RREQUISITN based on REQNSTAT details
     Set rq_grid_det = rdb.CreateQuery("grid_det", "select REQNNO,REQNSTAT,APNTDT,MEMONO,RECPTNO,MEMODT,EXAMDT,FREPORTDT,PREPORTDT ,FREPORTBY,RECPTDT,PATNSTAT,recptno,reqndt,ENTERBY from Rislib.RREQUISITN where REQNNO=? AND REQNSTAT='E'")
     rq_grid_det(0) = "aa"
     Set rs_grid_det = rq_grid_det.OpenResultset(2, 3)
     exam_flag = True
 
    ' STR = "SELECT * FROM RISLIB.RISREPORTS WHERE REQNNO=?"
    ' Set SELRISREPORT_RQ = rdb.CreateQuery("SELRISREPORT_RQ1", STR)
    ' SELRISREPORT_RQ(0) = "LZZCT01000000"
    ' Set SELRISREPORT_RS = SELRISREPORT_RQ.OpenResultset(2, 3)
     
'     STR = "SELECT * FROM RISLIB.RCLINIMPR WHERE REQNNO=?"
'     Set SELRCLINIMPR_RQ = rdb.CreateQuery("SELRCLINIMPR_RQ1 ", STR)
'     SELRCLINIMPR_RQ(0) = "LZZCT01000000"
'     Set SELRCLINIMPR_RS = SELRCLINIMPR_RQ.OpenResultset(2, 3)
'
'     STR = "SELECT * FROM RISLIB.RADDENDUM WHERE REQNO=?"
'     Set SELRADDENDUM_RQ = rdb.CreateQuery("SELRADDENDUM_RQ1", STR)
'     SELRADDENDUM_RQ(0) = "LZZCT01000000"
'     Set SELRADDENDUM_RS = SELRADDENDUM_RQ.OpenResultset(2, 3)
'     ADDFLAG = False
'     REPFLAG = False
'     IMPFLAG = False
'     STR = "SELECT * FROM RISLIB.RMODALITY WHERE MODALITYCD=?"
'     Set SELRMODALITY_QR = rdb.CreateQuery("SELRMODALITY1_QR1", STR)
'     SELRMODALITY_QR(0) = "CT"
'     Set SELRMODALITY_RS = SELRMODALITY_QR.OpenResultset(2, 3)
     
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
'    STR = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
'    Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO_QR111", STR)
'
'
'    ''REQNO SELECT FOR STAFF
'    STR = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO=?"
'    Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN1_QR1BC", STR)
'
'    ''DTMNO
'    STR = "SELECT * FROM DISLIB.DTMNO WHERE CASENO=?"
'    Set SELDTMNO_QR = rdb.CreateQuery("SELDTMNO_STR111", STR)
'
'    STR = "select * from Rislib.Rrequisitn where reqnNo=?"
'    Set selrrequisitnqr = rdb.CreateQuery("selrrequisitnqr1", STR)
'
    str = "select * from dislib.drequisitn where reqno=?"
    Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn1_qr1", str)
'
'    STR = "select * from patoolslib.ta1012 where case_no=?"
'    Set selta1012_qr = rdb.CreateQuery("1selta10121_qr1", STR)
'
'    STR = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
'    Set SELMA301_QR = rdb.CreateQuery("SELMA3011_QR1", STR)
'
'    STR = "SELECT * FROM DISLIB.DFREPORT WHERE REQNO=?"
'    Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT1_QR1", STR)
'
'    STR = "SELECT * FROM DISLIB.DADDENDUM WHERE REQNO=?"
'    Set SELDADDENDUMQR = rdb.CreateQuery("SELDADDENDUMQR1", STR)
'
    str = "SELECT * FROM DISLIB.DREQUISITN WHERE SPIHREQNO=?"
        Set SELDREQNOQR1 = rdb.CreateQuery("S1ELDREQNOQR121", str)
'
    str = "SELECT * FROM PATOOLSLIB.TA101 WHERE CAsE_NO=?"
    Set SELTA101_QR = rdb.CreateQuery("RSELTA101_QR1", str)
'
'    STR = "SELECT pro_code FROM DISLIB.DMULPROCE WHERE REQNO=? "
'    Set DMULPROCE_QR = rdb.CreateQuery("DMULPROCE_QR1TUMREPRQ", STR)
'
'    STR = "SELECT * FROM DISLIB.DSTREPORT WHERE REQNO=?"
'    Set SELSTREPORTQR = rdb.CreateQuery("SELSTREPORTQRSVZ", STR)
'
'    STR = "SELECT * FROM DISLIB.DECREPORT WHERE REQNO=?"
'    Set SELECREPORTQR = rdb.CreateQuery("SELDECREPORTQRXX1", STR)
'
'    STR = "SELECT * FROM DISLIB.DBMAREPORT WHERE REQNO=?"
'    Set SELDBMAREPORT_QR = rdb.CreateQuery("SELDBMAREPORT_STR", STR)
    
     Frame2.Enabled = True
     Opt_all.Enabled = False
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
Opt_all.Enabled = False
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
                             msf_reqstat.TextMatrix(srno, 2) = rs_allreq("caseno")
                             
                             
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
                                
                                SELDREQNOQR1(0) = rs_allreq("REQNNO")
                                Set SELDREQNORS1 = SELDREQNOQR1.OpenResultset(1, 3)
                                If SELDREQNORS1.RowCount > 0 Then
                                    msf_reqstat.CellBackColor = vbGreen
                                    msf_reqstat.TextMatrix(srno, 8) = req_chglen(SELDREQNORS1("REQNO"))
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

'Private Sub msf_reqstat_Click()
'On Error GoTo MSG:
''button_ena_dis ("print")
'Dim STR1 As String
'Dim STR2 As String
'Dim STR3 As String
'
'DoEvents
'
'msf_reqstat.Col = 1
'STR = usr_name & " " & ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
'VAR_REQNNO = req_chglen(Left(msf_reqstat.Text, 16))  '''req_chglen(mask_edit_req.Text)
'var_caseno = msf_reqstat.TextMatrix(msf_reqstat.Row, 2)
'selrrequisitnqr(0) = VAR_REQNNO
'Set selrrequisitnrs = selrrequisitnqr.OpenResultset(1, 3)
'If selrrequisitnrs.RowCount > 0 Then
'
'End If
'seldrequisitn_qr(0) = VAR_REQNNO
'Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'
'
''''DTM BLANK REPORT CHEQUE
'If selrrequisitnrs("MODALITYCD") = "SS" Then
'    If SELDTMTAB_FLAG = False Then
'        SELDTMTAB_QR(0) = VAR_REQNNO
'        Set SELDTMTAB_RS = SELDTMTAB_QR.OpenResultset(1, 3)
'        SELDTMTAB_FLAG = True
'    Else
'        SELDTMTAB_QR(0) = VAR_REQNNO
'        SELDTMTAB_RS.MoreResults
'        SELDTMTAB_RS.Requery
'    End If
'    'If ((Len(Trim(SELDTMTAB_RS("ABO"))) > 1 And SELDTMTAB_RS("ABO")) <> "") And (Len(Trim(SELDTMTAB_RS("RH"))) > 1 Or Trim(SELDTMTAB_RS("RH")) <> "") And (Len(Trim(SELDTMTAB_RS("ABS"))) > 1 Or Trim(SELDTMTAB_RS("ABS")) <> "") And (Len(Trim(SELDTMTAB_RS("IAT"))) > 1 Or Trim(SELDTMTAB_RS("IAT")) <> "") And (Len(Trim(SELDTMTAB_RS("DAT"))) > 1 Or Trim(SELDTMTAB_RS("DAT")) <> "") And (Len(Trim(SELDTMTAB_RS("HBSAG"))) > 1 Or Trim(SELDTMTAB_RS("HBSAG")) <> "") And (Len(Trim(SELDTMTAB_RS("HCV"))) > 1 Or Trim(SELDTMTAB_RS("HCV")) <> "") And (Len(Trim(SELDTMTAB_RS("HIV"))) > 1 Or Trim(SELDTMTAB_RS("HIV")) <> "") Then
'    If (IsNull(SELDTMTAB_RS("ABO")) Or Trim(SELDTMTAB_RS("ABO")) = "") And (IsNull(SELDTMTAB_RS("RH")) Or Trim(SELDTMTAB_RS("RH")) = "") And (IsNull(Trim(SELDTMTAB_RS("ABS"))) Or Trim(SELDTMTAB_RS("ABS")) = "") And (IsNull(Trim(SELDTMTAB_RS("IAT"))) Or Trim(SELDTMTAB_RS("IAT")) = "") And (IsNull(Trim(SELDTMTAB_RS("DAT"))) Or Trim(SELDTMTAB_RS("DAT")) = "") And (IsNull(Trim(SELDTMTAB_RS("HBSAG"))) Or Trim(SELDTMTAB_RS("HBSAG")) = "") And (IsNull(Trim(SELDTMTAB_RS("HCV"))) Or Trim(SELDTMTAB_RS("HCV")) = "") And (IsNull(Trim(SELDTMTAB_RS("HIV"))) Or Trim(SELDTMTAB_RS("HIV")) = "") Then
'    'If (IsNull(SELDTMTAB_RS("ABS")) Or Trim(SELDTMTAB_RS("ABS")) = "") And (IsNull(Trim(SELDTMTAB_RS("IAT"))) Or Trim(SELDTMTAB_RS("IAT")) = "") Then
'        MsgBox " Provisional has not Created !!!", vbInformation, "DTM"
'        Exit Sub
'    Else
'        GoTo DTM
'    End If
'End If
''''DTM BLANK REPORT CHEQUE
'DTM:
'For i = 0 To 25
'    CRYSTALREPORT1.Formulas(i) = ""
'Next i
'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"   'AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'
'If Trim(msf_reqstat.TextMatrix(msf_reqstat.Row, 7)) = "" Then
'    MsgBox "Provisional Report is Not yet Prepared"
'    Exit Sub
'End If
'If SELTA101_FLAG = False Then
'    SELTA101_QR(0) = selrrequisitnrs("CASENO")
'    Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
'    SELTA101_FLAG = True
'Else
'    SELTA101_QR(0) = selrrequisitnrs("CASENO")
'    SELTA101_RS.MoreResults
'    SELTA101_RS.Requery
'End If
'If SELEMPBASINFOFLAG = False Then
'    SELEMPBASINFO_QR(0) = msk_caseno.Text
'    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
'    SELEMPBASINFOFLAG = True
'Else
'    SELEMPBASINFO_QR(0) = msk_caseno.Text
'    SELEMPBASINFO_RS.MoreResults
'    SELEMPBASINFO_RS.Requery
'End If
'
'selta1012_qr(0) = selrrequisitnrs("CASENO")
'Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
'
'STR = usr_name & " " & ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'
'
''*** SURGICAL PAHOLOGY REPORT
'''If selrrequisitnrs("modalitycd") = "SP" Then
'''
'''        If selrrequisitnrs("COMMITFLAG") = "Y" Then
'''           SELDADDENDUMQR(0) = VAR_REQNNO
'''           Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
'''            If SELDADDENDUMRS.RowCount > 0 Then
'''                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISSPADD.RPT"
'''            Else
'''                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISSP.RPT"
'''            End If
'''        Else
'''            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISSP.RPT"
'''        End If
'''        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'''        seldrequisitn_qr(0) = VAR_REQNNO
'''        Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'''        If seldrequisitn_rs.RowCount > 0 Then
'''             CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
'''        Else
'''            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'''        End If
'''        selta1012_qr(0) = selrrequisitnrs("CASENO")
'''        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
'''        If selta1012_rs.RowCount > 0 Then
'''            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
'''        Else
'''            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & "" & "'"
'''        End If
'''
'''
'''        SELMA301_QR(0) = seldrequisitn_rs("sITE")
'''        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
'''        If SELMA301_RS.RowCount > 0 Then
'''            CRYSTALREPORT1.Formulas(3) = "SIT = '" & Trim(SELMA301_RS("SITEDESC")) & "'"
'''        Else
'''            CRYSTALREPORT1.Formulas(3) = "SIT = '" & "" & "'"
'''        End If
'''        If seldrequisitn_rs.RowCount > 0 Then
'''            CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & Trim(seldrequisitn_rs("OPPROCEDUR")) & "'"
'''        Else
'''            CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & "" & "'"
'''        End If
'''
'''        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'''            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "PROVISIONAL REPORT" & "'"
'''        Else
'''            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "FINAL REPORT" & "'"
'''        End If
'''        SELDFREPORT_QR(0) = VAR_REQNNO
'''        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
'''        If SELDFREPORT_RS.RowCount > 0 Then
'''            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
'''                CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "FROZEN SECTION" & "'"
'''            Else
'''                CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "" & "'"
'''            End If
'''        Else
'''            CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "" & "'"
'''        End If
'''        If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
'''            CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
'''        Else
'''            CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "" & "'"
'''        End If
'''        CRYSTALREPORT1.Formulas(8) = "USRID = '" & STR & "'"
'''        CRYSTALREPORT1.Action = 1
''' End If
'
'selta1012_qr(0) = selrrequisitnrs("CASENO")
'Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
'
'
'
''''********BONE MARROW*************
'
'If selrrequisitnrs("MODALITYCD") = "BM" Then
'    If SELDBMAREPORT_FLAG = False Then
'        SELDBMAREPORT_QR(0) = VAR_REQNNO
'        Set SELDBMAREPORT_RS = SELDBMAREPORT_QR.OpenResultset(1, 3)
'        SELDBMAREPORT_FLAG = True
'    Else
'        SELDBMAREPORT_QR(0) = VAR_REQNNO
'        SELDBMAREPORT_RS.MoreResults
'        SELDBMAREPORT_RS.Requery
'    End If
'    If opt_patient.Value = True Then
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISBM1.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Not selta1012_rs.EOF Then
'            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & selta1012_rs("PRVDIAGDES") & "'"
'        End If
'        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
'            STR1 = ""
'            STR1 = "Entered By " & selrrequisitnrs("ENTERBY")
'            If Len(Trim(selrrequisitnrs("ENTERBY"))) > 0 Then
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            Else
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & "" & "'"
'            End If
'         Else
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BONE MARROW ASPIRATION REPORT" & "'"
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'        If Trim(SELDBMAREPORT_RS("CD2")) = "" And Trim(SELDBMAREPORT_RS("CD3")) = "" And Trim(SELDBMAREPORT_RS("CD4")) = "" And Trim(SELDBMAREPORT_RS("CD5")) = "" And Trim(SELDBMAREPORT_RS("CD7")) = "" And Trim(SELDBMAREPORT_RS("CD8")) = "" And Trim(SELDBMAREPORT_RS("CYTOCD3")) = "" Then
'            CRYSTALREPORT1.Formulas(13) = "tmarkers = '" & "" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(13) = "tmarkers = '" & "abc" & "'"
'        End If
'        If Trim(SELDBMAREPORT_RS("CD10")) = "" And Trim(SELDBMAREPORT_RS("CD19")) = "" And Trim(SELDBMAREPORT_RS("FMC7")) = "" And Trim(SELDBMAREPORT_RS("CD22")) = "" And Trim(SELDBMAREPORT_RS("CD23")) = "" And Trim(SELDBMAREPORT_RS("smig")) = "" And Trim(SELDBMAREPORT_RS("CD79B")) = "" And Trim(SELDBMAREPORT_RS("CD25")) = "" And Trim(SELDBMAREPORT_RS("CD11C")) = "" And Trim(SELDBMAREPORT_RS("cd103")) = "" And Trim(SELDBMAREPORT_RS("kappa")) = "" And Trim(SELDBMAREPORT_RS("lamda")) = "" And Trim(SELDBMAREPORT_RS("DBA44")) = "" And Trim(SELDBMAREPORT_RS("cytoEP22")) = "" And Trim(SELDBMAREPORT_RS("cytocd79a")) = "" And Trim(SELDBMAREPORT_RS("CYTOMU")) = "" Then
'            CRYSTALREPORT1.Formulas(14) = "bmarkers = '" & "" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(14) = "bmarkers = '" & "abc" & "'"
'        End If
'        If Trim(SELDBMAREPORT_RS("CD34")) = "" And Trim(SELDBMAREPORT_RS("HLADR")) = "" And Trim(SELDBMAREPORT_RS("CD38")) = "" And Trim(SELDBMAREPORT_RS("TDT")) = "" And Trim(SELDBMAREPORT_RS("CD138")) = "" And Trim(SELDBMAREPORT_RS("ANTIMPO")) = "" Then
'            CRYSTALREPORT1.Formulas(15) = "others = '" & "" & "'"
'        Else
'                CRYSTALREPORT1.Formulas(15) = "others = '" & "abc" & "'"
'        End If
'        If Trim(SELDBMAREPORT_RS("cd13")) = "" And Trim(SELDBMAREPORT_RS("Cd14")) = "" And Trim(SELDBMAREPORT_RS("CD15")) = "" And Trim(SELDBMAREPORT_RS("CD33")) = "" And Trim(SELDBMAREPORT_RS("cd36")) = "" And Trim(SELDBMAREPORT_RS("cd4142")) = "" And Trim(SELDBMAREPORT_RS("Glycophoin")) = "" And Trim(SELDBMAREPORT_RS("CD68")) = "" And Trim(SELDBMAREPORT_RS("CD42")) = "" And Trim(SELDBMAREPORT_RS("CD117")) = "" Then
'           CRYSTALREPORT1.Formulas(16) = "mm = '" & "" & "'"
'        Else
'           CRYSTALREPORT1.Formulas(16) = "mm = '" & "abc" & "'"
'        End If
'        If Trim(SELDBMAREPORT_RS("IRONSTAIN")) = "" And Trim(SELDBMAREPORT_RS("GIDEROBLAS")) = "" And Trim(SELDBMAREPORT_RS("sudanblack")) = "" And Trim(SELDBMAREPORT_RS("TOLUIDINEB")) = "" Then
'           CRYSTALREPORT1.Formulas(17) = "cci = '" & "" & "'"
'        Else
'           CRYSTALREPORT1.Formulas(17) = "cci = '" & "abc" & "'"
'        End If
'        If Trim(SELDBMAREPORT_RS("myelopero")) = "" And Trim(SELDBMAREPORT_RS("nonspester")) = "" And Trim(SELDBMAREPORT_RS("trap")) = "" Then
'           CRYSTALREPORT1.Formulas(18) = "ccm = '" & "" & "'"
'        Else
'           CRYSTALREPORT1.Formulas(18) = "ccm = '" & "abc" & "'"
'        End If
'        If Trim(SELDBMAREPORT_RS("IMPRINT")) = "" And Trim(SELDBMAREPORT_RS("abcellblas")) = "" Then
'           CRYSTALREPORT1.Formulas(19) = "im = '" & "" & "'"
'        Else
'           CRYSTALREPORT1.Formulas(19) = "im = '" & "abc" & "'"
'        End If
'        If Trim(SELDBMAREPORT_RS("mcellulari")) = "" And Trim(SELDBMAREPORT_RS("morphology")) = "" Then
'           CRYSTALREPORT1.Formulas(20) = "ms = '" & "" & "'"
'        Else
'           CRYSTALREPORT1.Formulas(20) = "ms = '" & "abc" & "'"
'        End If
'        If Trim(SELDBMAREPORT_RS("ecellulari")) = "" And Trim(SELDBMAREPORT_RS("maturation")) = "" Then
'           CRYSTALREPORT1.Formulas(21) = "es = '" & "" & "'"
'        Else
'           CRYSTALREPORT1.Formulas(21) = "es = '" & "abc" & "'"
'        End If
'        If Trim(SELDBMAREPORT_RS("mature")) = "" And Trim(SELDBMAREPORT_RS("immature")) = "" And Trim(SELDBMAREPORT_RS("atypbinu")) = "" And Trim(SELDBMAREPORT_RS("plasmablas")) = "" Then
'           CRYSTALREPORT1.Formulas(22) = "pc = '" & "" & "'"
'        Else
'           CRYSTALREPORT1.Formulas(22) = "pc = '" & "abc" & "'"
'        End If
'        If Trim(SELDBMAREPORT_RS("MLYMPHOCYT")) = "" And Trim(SELDBMAREPORT_RS("ALYMPHOCYT")) = "" And Trim(SELDBMAREPORT_RS("Plasmacell")) = "" Then
'           CRYSTALREPORT1.Formulas(23) = "ls = '" & "" & "'"
'        Else
'           CRYSTALREPORT1.Formulas(23) = "ls = '" & "abc" & "'"
'        End If
'        If Trim(SELDBMAREPORT_RS("PROMYELOCY")) = "" And Trim(SELDBMAREPORT_RS("PROMYELOCY")) = "" And Trim(SELDBMAREPORT_RS("MMYELOCYTE")) = "" And Trim(SELDBMAREPORT_RS("polymorphs")) = "" And Trim(SELDBMAREPORT_RS("basophils")) = "" And Trim(SELDBMAREPORT_RS("eosinophil")) = "" And Trim(SELDBMAREPORT_RS("monocytes")) = "" And Trim(SELDBMAREPORT_RS("PROMONOCYT")) = "" Then
'           CRYSTALREPORT1.Formulas(24) = "mys = '" & "" & "'"
'        Else
'           CRYSTALREPORT1.Formulas(24) = "mys = '" & "abc" & "'"
'        End If
'        If Trim(SELDBMAREPORT_RS("ECELLULARI")) = "" And Trim(SELDBMAREPORT_RS("MERATIO")) = "" Then
'           CRYSTALREPORT1.Formulas(25) = "mor = '" & "" & "'"
'        Else
'           CRYSTALREPORT1.Formulas(25) = "mor = '" & "abc" & "'"
'        End If
'    End If
'    CRYSTALREPORT1.WindowState = crptMaximized
'    CRYSTALREPORT1.Action = 1
'End If
'
'
'''******BLOOD PRODUCT ****************
'
'If selrrequisitnrs("MODALITYCD") = "BP" Then
'    If SELDTMNO_FLAG = False Then
'        SELDTMNO_QR(0) = selrrequisitnrs("CASENO")
'        Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
'        SELDTMNO_FLAG = True
'    Else
'        SELDTMNO_QR(0) = selrrequisitnrs("CASENO")
'        SELDTMNO_RS.MoreResults
'        SELDTMNO_RS.Requery
'    End If
'    If opt_patient.Value = True Then
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISBP.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        CRYSTALREPORT1.Formulas(1) = "MOD = '" & "BLOOD PRODUCT" & "'"
'        If selrrequisitnrs("PATNSTAT") = "OP" Then
'            CRYSTALREPORT1.Formulas(2) = "ST = '" & selrrequisitnrs("pat_catg") & "/" & " Out Patient" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(2) = "ST = '" & selrrequisitnrs("pat_catg") & "/" & " In patient" & "'"
'        End If
'        CRYSTALREPORT1.Formulas(3) = "SEXAGE = '" & SELTA101_RS("SEX") & "/" & DateDiff("yyyy", SELTA101_RS("BIRTH_DATE"), selrrequisitnrs("REQNDT")) & "'"
'        If selrrequisitnrs("PATNSTAT") = "IP" Then
'            CRYSTALREPORT1.Formulas(4) = "WRBED = '" & seldrequisitn_rs("WARD_NAME") & "/" & seldrequisitn_rs("bedno") & "'"
'        Else
'            CRYSTALREPORT1.Formulas(4) = "WRBED = '" & " " & "'"
'        End If
'        CRYSTALREPORT1.Formulas(5) = "PATHNO = '" & SELDTMNO_RS("DTMNO") & "'"
'        CRYSTALREPORT1.Formulas(6) = "USRID = '" & STR & "'"
'    Else
'
'    End If
'    CRYSTALREPORT1.WindowState = crptMaximized
'    CRYSTALREPORT1.Action = 1
'End If
'''******BLOOD PRODUCT ****************
'''******SCREENING TEST ****************
'If selrrequisitnrs("MODALITYCD") = "SS" Then
'    If opt_patient.Value = True Then
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISdtm.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If seldrequisitn_rs.RowCount > 0 Then
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
'        Else
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'        End If
'        CRYSTALREPORT1.Formulas(2) = "USRID = '" & STR & "'"
'    Else
'        CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISdtmreq.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If seldrequisitn_rs.RowCount > 0 Then
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
'        Else
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'        End If
'        CRYSTALREPORT1.Formulas(2) = "USRID = '" & STR & "'"
'        If SELEMPBASINFO_RS.RowCount > 0 Then
'            CRYSTALREPORT1.Formulas(3) = "NAM='" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'        Else
'            CRYSTALREPORT1.Formulas(3) = "NAM= '" & " " & "'"
'        End If
'        CRYSTALREPORT1.Formulas(4) = "STDOB='" & SELEMPBASINFO_RS("GENDER") & "'"
'        CRYSTALREPORT1.Formulas(5) = "DOB='" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT")) & "'"
'    End If
'    CRYSTALREPORT1.WindowState = crptMaximized
'    CRYSTALREPORT1.Action = 1
'End If
'''*****SCREENING TEST ****************
'''*******HLA TYPING *****************
'If selrrequisitnrs("MODALITYCD") = "HT" Then
'    If opt_patient.Value = True Then
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISHLA.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        CRYSTALREPORT1.Formulas(7) = "USRID = '" & STR & "'"
'        If Len(Trim(seldrequisitn_rs("LABNO"))) > 0 Then
'            CRYSTALREPORT1.Formulas(8) = "NAM = '" & seldrequisitn_rs("NMETERIALS") & "'"
'            CRYSTALREPORT1.Formulas(9) = "RELATION = 'Relation' "
'        Else
'            CRYSTALREPORT1.Formulas(8) = "NAM = '" & Trim(SELTA101_RS("FIRST_NAME")) + " " + Trim(SELTA101_RS("MID_NAME")) + " " + Trim(SELTA101_RS("LAST_NAME")) & "'"
'            CRYSTALREPORT1.Formulas(9) = "RELATION ='" & " " & "'"
'        End If
'        CRYSTALREPORT1.Formulas(9) = "dtmno='" & " " & "'"
'        CRYSTALREPORT1.WindowState = crptMaximized
'        CRYSTALREPORT1.Action = 1
'    Else
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISHLAstaff.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        CRYSTALREPORT1.Formulas(7) = "USRID = '" & STR & "'"
'        If Len(Trim(seldrequisitn_rs("LABNO"))) > 0 Then
'            CRYSTALREPORT1.Formulas(8) = "NAM = '" & seldrequisitn_rs("NMETERIALS") & "'"
'            CRYSTALREPORT1.Formulas(9) = "RELATION = 'Relation' "
'        Else
'            CRYSTALREPORT1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'            CRYSTALREPORT1.Formulas(9) = "RELATION ='" & " " & "'"
'        End If
'        CRYSTALREPORT1.Formulas(10) = "STDOB ='" & SELEMPBASINFO_RS("GENDER") & "'"
'        CRYSTALREPORT1.Formulas(11) = "DOB ='" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT")) & "'"
'        CRYSTALREPORT1.WindowState = crptMaximized
'        CRYSTALREPORT1.Action = 1
'    End If
'End If
'''*******HLA TYPING *****************
' '*** SURGICAL PAHOLOGY REPORT
' '*** IMMUNIOHISTO CHEMISTRY REPORT
' If selrrequisitnrs("MODALITYCD") = "IH" Then
'    If opt_patient.Value = True Then
'        If Len(Trim(seldrequisitn_rs("SPIHREQNO"))) = 0 Then
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISIH.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            If selrrequisitnrs.RowCount > 0 Then
'               If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'                    CRYSTALREPORT1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHC REPORT" & "'"
'                    'CRYSTALREPORT1.Formulas(3) = "ENT = '" & "ENTERED BY " & cmb_enterby & "'"
'                Else
'                    CRYSTALREPORT1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHC REPORT" & "'"
'                End If
'            End If
'            If selta1012_rs.RowCount > 0 Then
'                CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
'            Else
'                CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & "" & "'"
'            End If
'            If seldrequisitn_rs.RowCount > 0 Then
'                CRYSTALREPORT1.Formulas(3) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
'            Else
'                CRYSTALREPORT1.Formulas(3) = "PREPATH = '" & "" & "'"
'            End If
'            CRYSTALREPORT1.Formulas(4) = "USRID = '" & STR & "'"
'            CRYSTALREPORT1.WindowState = crptMaximized
'            CRYSTALREPORT1.Action = 1
'        End If
'    Else
'        If Len(Trim(seldrequisitn_rs("SPIHREQNO"))) = 0 Then
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISIHSTAFF.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            If selrrequisitnrs.RowCount > 0 Then
'               If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'                    CRYSTALREPORT1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHC REPORT" & "'"
'                    'CRYSTALREPORT1.Formulas(3) = "ENT = '" & "ENTERED BY " & cmb_enterby & "'"
'                Else
'                    CRYSTALREPORT1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHC REPORT" & "'"
'                End If
'            End If
'            If selta1012_rs.RowCount > 0 Then
'                CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
'            Else
'                CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & "" & "'"
'            End If
'            If seldrequisitn_rs.RowCount > 0 Then
'                CRYSTALREPORT1.Formulas(3) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
'            Else
'                CRYSTALREPORT1.Formulas(3) = "PREPATH = '" & "" & "'"
'            End If
'            CRYSTALREPORT1.Formulas(4) = "USRID = '" & STR & "'"
'            CRYSTALREPORT1.Formulas(5) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'            CRYSTALREPORT1.Formulas(6) = "SEX = '" & SELEMPBASINFO_RS("GENDER") & "'"
'            CRYSTALREPORT1.Formulas(4) = "SIT = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
'            CRYSTALREPORT1.WindowState = crptMaximized
'            CRYSTALREPORT1.Action = 1
'        End If
'    End If
'  ' *** IMMUNIO HISTO CHEMISTRY REPORT
'End If
'
'If selrrequisitnrs("MODALITYCD") = "SP" Then
'  ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
'    CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'    If opt_patient.Value = True Then
'
'        If selrrequisitnrs("COMMITFLAG") = "Y" Then
'           SELDADDENDUMQR(0) = VAR_REQNNO
'           Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
'            If SELDADDENDUMRS.RowCount > 0 Then
'                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
'            Else
'                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
'            End If
'        Else
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
'        End If
'        seldrequisitn_qr(0) = VAR_REQNNO
'        Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'        If seldrequisitn_rs.RowCount > 0 Then
'             CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
'        Else
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'        End If
'        selta1012_qr(0) = selrrequisitnrs("CASENO")
'        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
'        If selta1012_rs.RowCount > 0 Then
'            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
'        Else
'            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & "" & "'"
'        End If
'
'        Dim Z
'        Dim Y
'        SELMA301_QR(0) = seldrequisitn_rs("SITE")
'        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
'        If SELMA301_RS.RowCount > 0 Then
'            CRYSTALREPORT1.Formulas(3) = "SIT = '" & Trim(SELMA301_RS("SITEDESC")) & "'"
'        Else
'            CRYSTALREPORT1.Formulas(3) = "SIT = '" & "" & "'"
'        End If
'        If seldrequisitn_rs.RowCount > 0 Then
'            If Trim(seldrequisitn_rs("OPPROCEDUR")) <> "" Then
'                Z = InStr(1, Trim(seldrequisitn_rs("OPPROCEDUR")), "'")
'                If Z > 0 Then
'                    Y = Trim(seldrequisitn_rs("OPPROCEDUR"))
'                    U = Len(Y)
'                    x = Left(Y, Z - 1)
'                    w = Right(Y, U - (Z))
'                    Y = x & w
'                    CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & Y & "'"
'                Else
'                    CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & Trim(seldrequisitn_rs("OPPROCEDUR")) & "'"
'                End If
'            Else
'                CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & Trim(seldrequisitn_rs("OPPROCEDUR")) & "'"
'            End If
'        Else
'            CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & "" & "'"
'        End If
'        STR1 = ""
'        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
'            STR1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        Else
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        SELDFREPORT_QR(0) = VAR_REQNNO
'        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
'        If SELDFREPORT_RS.RowCount > 0 Then
'            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
'                CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "" & "'"
'            End If
'            If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
'                CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "" & "'"
'            End If
'            If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
'                CRYSTALREPORT1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(10) = "GROSDESC = '" & "" & "'"
'            End If
'            If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
'                CRYSTALREPORT1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(11) = "MICRODESC = '" & "" & "'"
'            End If
'            If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
'                CRYSTALREPORT1.Formulas(12) = "IMPN = '" & "Impression" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(12) = "IMPN = '" & "" & "'"
'            End If
'
'        End If
'
'        CRYSTALREPORT1.Formulas(8) = "USRID = '" & STR & "'"
'
'        'CrystalReport1.Formulas(10) = "GROSDESC = '" & "Gross description" & "'"
'        'CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
'        CRYSTALREPORT1.WindowState = crptMaximized
'        CRYSTALREPORT1.Action = 1
'    Else
'        If selrrequisitnrs("COMMITFLAG") = "Y" Then
'            SELDADDENDUMQR(0) = VAR_REQNNO
'            Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
'            If SELDADDENDUMRS.RowCount > 0 Then
'                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENADDSTAFF.RPT"
'            Else
'                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
'            End If
'        Else
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENSTAFF.RPT"
'        End If
'        seldrequisitn_qr(0) = VAR_REQNNO
'        Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'        If seldrequisitn_rs.RowCount > 0 Then
'             CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
'        Else
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'        End If
'        selta1012_qr(0) = selrrequisitnrs("CASENO")
'        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
'        If selta1012_rs.RowCount > 0 Then
'            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
'        Else
'            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & "" & "'"
'        End If
'
'
'        SELMA301_QR(0) = seldrequisitn_rs("SITE")
'        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
'        If SELMA301_RS.RowCount > 0 Then
'            CRYSTALREPORT1.Formulas(3) = "SIT = '" & Trim(SELMA301_RS("SITEDESC")) & "'"
'        Else
'            CRYSTALREPORT1.Formulas(3) = "SIT = '" & "" & "'"
'        End If
'        If seldrequisitn_rs.RowCount > 0 Then
'            CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & Trim(seldrequisitn_rs("OPPROCEDUR")) & "'"
'        Else
'            CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & "" & "'"
'        End If
'        STR1 = ""
'        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
'            STR1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        Else
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        SELDFREPORT_QR(0) = VAR_REQNNO
'        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
'        If SELDFREPORT_RS.RowCount > 0 Then
'            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
'                CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "" & "'"
'            End If
'            If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
'                CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "" & "'"
'            End If
'            If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
'                CRYSTALREPORT1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(10) = "GROSDESC = '" & "" & "'"
'            End If
'            If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
'                CRYSTALREPORT1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(11) = "MICRODESC = '" & "" & "'"
'            End If
'            If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
'                CRYSTALREPORT1.Formulas(12) = "IMPN = '" & "Impression" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(12) = "IMPN = '" & "" & "'"
'            End If
'
'        End If
'
'        CRYSTALREPORT1.Formulas(8) = "USRID = '" & STR & "'"
'        CRYSTALREPORT1.Formulas(13) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'        CRYSTALREPORT1.Formulas(14) = "DOB = '" & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT")) & "'"
'        CRYSTALREPORT1.Formulas(15) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
'        CRYSTALREPORT1.WindowState = crptMaximized
'        CRYSTALREPORT1.Action = 1
'    End If
' ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
' End If
' '**** FNAC REPORT
'If selrrequisitnrs("MODALITYCD") = "FA" Then
'    If opt_patient.Value = True Then
'        If selrrequisitnrs("REQNSTAT") <> "F" Then
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISFA.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
'                STR1 = ""
'                STR1 = "Entered By " & selrrequisitnrs("ENTERBY")
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'             Else
'
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            End If
'            CRYSTALREPORT1.Formulas(7) = "USRID = '" & STR & "'"
'        Else
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISFAADD.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            CRYSTALREPORT1.Formulas(7) = "USRID = '" & STR & "'"
'        End If
'        CRYSTALREPORT1.WindowState = crptMaximized
'        CRYSTALREPORT1.Action = 1
'    Else
'        If selrrequisitnrs("REQNSTAT") <> "F" Then
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISFASTAFF.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
'                STR1 = ""
'                STR1 = "Entered By " & selrrequisitnrs("ENTERBY")
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'             Else
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "FINE NEEDLE ASPIRATION CYTOLOGY REPORT" & "'"
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            End If
'            CRYSTALREPORT1.Formulas(7) = "USRID = '" & STR & "'"
'        Else
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISFAADDSTAFF.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            CRYSTALREPORT1.Formulas(7) = "USRID = '" & STR & "'"
'        End If
'        CRYSTALREPORT1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'        CRYSTALREPORT1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
'        CRYSTALREPORT1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
'        CRYSTALREPORT1.WindowState = crptMaximized
'        CRYSTALREPORT1.Action = 1
'    End If
'End If
' '**** FNAC REPORT
' '**** CYTOLOGY REPORT
'If selrrequisitnrs("MODALITYCD") = "CG" Or selrrequisitnrs("MODALITYCD") = "CN" Then
'    If opt_patient.Value = True Then
'        If selrrequisitnrs("REQNSTAT") <> "F" Then
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISCG.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
''            If Trim(TXT_CSITE.Text) = "" Then
''                CrystalReport1.Formulas(1) = "STL = '" & "" & "'"
''            Else
''                CrystalReport1.Formulas(1) = "STL = '" & "Site  " & TXT_CSITE.Text & "'"
''            End If
'            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
'                STR1 = ""
'                STR1 = "Entered By " & Trim(selrrequisitnrs("ENTERBY"))
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'             Else
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            End If
'            CRYSTALREPORT1.Formulas(7) = "USRID = '" & STR & "'"
'            CRYSTALREPORT1.Formulas(12) = "NOSLIDES=" & "" & seldrequisitn_rs("LABNO")
'        Else
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISCGADD.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            CRYSTALREPORT1.Formulas(7) = "USRID = '" & STR & "'"
'            CRYSTALREPORT1.Formulas(12) = "NOSLIDES = " & "" & seldrequisitn_rs("LABNO")
'        End If
'        If Trim(selrrequisitnrs("PREPORTBY")) <> "" And Trim(selrrequisitnrs("FREPORTBY")) <> "" And UCase(Trim(selrrequisitnrs("PREPORTBY"))) <> UCase(Trim(selrrequisitnrs("FREPORTBY"))) Then
'            CRYSTALREPORT1.Formulas(11) = "P= '" & Trim(selrrequisitnrs("PREPORTBY")) & "'"
'            CRYSTALREPORT1.Formulas(12) = "F= '" & Trim(selrrequisitnrs("FREPORTBY")) & "'"
'            CRYSTALREPORT1.Formulas(13) = "PCAP= '" & "Cytotechnologist" & " '"
'            CRYSTALREPORT1.Formulas(14) = "FCAP= '" & "Pathologist" & " '"
'        End If
'        If Trim(selrrequisitnrs("PREPORTBY")) <> "" And Trim(selrrequisitnrs("FREPORTBY")) = "" Then
'            CRYSTALREPORT1.Formulas(11) = "PONLY= '" & Trim(selrrequisitnrs("PREPORTBY")) & "'"
'            'CrystalReport1.Formulas(12) = "F= '" & "" & "'"
'            CRYSTALREPORT1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
'        End If
'        If Trim(selrrequisitnrs("PREPORTBY")) <> "" And Trim(selrrequisitnrs("FREPORTBY")) <> "" And UCase(Trim(selrrequisitnrs("PREPORTBY"))) = UCase(Trim(selrrequisitnrs("FREPORTBY"))) Then
'            CRYSTALREPORT1.Formulas(11) = "PONLY= '" & Trim(selrrequisitnrs("PREPORTBY")) & "'"
'            CRYSTALREPORT1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
'        End If
'    Else
'        If selrrequisitnrs("REQNSTAT") <> "F" Then
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISCGSTAFF.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
''            If Trim(TXT_CSITE.Text) = "" Then
''                CrystalReport1.Formulas(1) = "STL = '" & "" & "'"
''            Else
''                CrystalReport1.Formulas(1) = "STL = '" & "Site  " & TXT_CSITE.Text & "'"
''            End If
'            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
'                STR1 = ""
'                STR1 = "Entered By " & Trim(selrrequisitnrs("ENTERBY"))
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'             Else
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "CYTOLOGY REPORT" & "'"
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            End If
'            CRYSTALREPORT1.Formulas(7) = "USRID = '" & STR & "'"
'        Else
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISCGADDSTAFF.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            CRYSTALREPORT1.Formulas(7) = "USRID = '" & STR & "'"
'        End If
'        CRYSTALREPORT1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'        CRYSTALREPORT1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
'        CRYSTALREPORT1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
'    End If
'    CRYSTALREPORT1.WindowState = crptMaximized
'    CRYSTALREPORT1.Action = 1
' End If
' '**** CYTOLOGY REPORT
' '**** BIOCHEMISTRY REPORT
' If selrrequisitnrs("MODALITYCD") = "BC" Then
'    If opt_patient.Value = True Then
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISBC.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Trim(prepath) = "" Then
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
'        'CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
'        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            STR1 = ""
'            STR1 = "Entered By " & selrrequisitnrs("enterby")
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'         Else
'
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
''        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
''        If pcount <> 0 Then
''            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
''        Else
'        CRYSTALREPORT1.Formulas(13) = "PAC = '" & "" & "'"
''        End If
'    Else
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISBCSTAFF.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Trim(prepath) = "" Then
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'        End If
'        'CrystalReport1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
'        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            STR1 = ""
'            STR1 = "Entered By " & selrrequisitnrs("enterby")
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'         Else
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
''        CrystalReport1.Formulas(12) = "USRID = '" & STR & "'"
''        If pcount <> 0 Then
''            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
''        Else
'        CRYSTALREPORT1.Formulas(13) = "PAC = '" & "" & "'"
''        End If
'    CRYSTALREPORT1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'    CRYSTALREPORT1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
'    CRYSTALREPORT1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
'    End If
'    CRYSTALREPORT1.WindowState = crptMaximized
'    CRYSTALREPORT1.Action = 1
' End If
' '**** BIOCHEMISTRY REPORT
' '**** EMERGENCY LAB REPORT
' If selrrequisitnrs("MODALITYCD") = "EL" Then
'    If opt_patient.Value = True Then
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISEL.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Trim(prepath) = "" Then
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
'        End If
'        CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
'        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            STR1 = ""
'            STR1 = "Entered By " & selrrequisitnrs("ENTERBY")
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'         Else
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'        If pcount <> 0 Then
'            CRYSTALREPORT1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(13) = "PAC = '" & "" & "'"
'        End If
'    Else
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISELSTAFF.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'        If Trim(prepath) = "" Then
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
'        End If
'        CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & txt_procedure & "'"
'        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            STR1 = ""
'            STR1 = "Entered By " & selrrequisitnrs("ENTERBY")
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'         Else
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'        If pcount <> 0 Then
'            CRYSTALREPORT1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(13) = "PAC = '" & "" & "'"
'        End If
'         CRYSTALREPORT1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'        CRYSTALREPORT1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
'        CRYSTALREPORT1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
'    End If
'    CRYSTALREPORT1.WindowState = crptMaximized
'    CRYSTALREPORT1.Action = 1
' End If
' '**** EMERGENCY LAB REPORT
' '**** TUMOUR MARKER REPORT
' If selrrequisitnrs("MODALITYCD") = "TM" Then
'     If DMULPROCEFLAG = False Then
'         DMULPROCE_QR(0) = VAR_REQNNO
'         Set DMULPROCE_RS = DMULPROCE_QR.OpenResultset(1, 3)
'         DMULPROCEFLAG = True
'     Else
'         DMULPROCE_QR(0) = VAR_REQNNO
'         DMULPROCE_RS.MoreResults
'         DMULPROCE_RS.Requery
'     End If
'     If DMULPROCE_RS.RowCount > 0 Then
'         While Not DMULPROCE_RS.EOF
'             If Trim(DMULPROCE_RS("PRO_CODE")) = "F816" Then
'                 IMFLAG = True
'             Else
'                 IMFLAG1 = True
'             End If
'             DMULPROCE_RS.MoveNext
'         Wend
'     End If
'     If opt_patient.Value = True Then
'         If IMFLAG1 = True Then
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISTM.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            If Trim(prepath) = "" Then
'                CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
'            End If
'            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
'                STR1 = ""
'                STR1 = "Entered by " & selrrequisitnrs("ENTERBY")
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            Else
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            End If
'            CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'        End If
'        If IMFLAG = True Then
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISTM1.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            If Trim(prepath) = "" Then
'                CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
'            End If
'            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
'                STR1 = ""
'                STR1 = "Entered by " & selrrequisitnrs("ENTERBY")
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            Else
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            End If
'            CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'        End If
'    Else
'        If IMFLAG1 = True Then
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISTMSTAFF.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            If Trim(prepath) = "" Then
'                CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
'            End If
'            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
'                STR1 = ""
'                STR1 = "Entered by " & selrrequisitnrs("ENTERBY")
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'             Else
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            End If
'            CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'            CRYSTALREPORT1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'            CRYSTALREPORT1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
'            CRYSTALREPORT1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
'        End If
'        If IMFLAG = True Then
'            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISTM1STAFF.RPT"
'            CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            If Trim(prepath) = "" Then
'                CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "Previous Path No." & Left(prepath, 100) & "'"
'            End If
'            If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
'                STR1 = ""
'                STR1 = "Entered by " & selrrequisitnrs("ENTERBY")
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            Else
'                CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "TUMOUR MARKER REPORT" & "'"
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            End If
'            CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'            CRYSTALREPORT1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'            CRYSTALREPORT1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
'            CRYSTALREPORT1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
'        End If
'    End If
'    CRYSTALREPORT1.WindowState = crptMaximized
'    CRYSTALREPORT1.Action = 1
'End If
''
' IMFLAG = False
' IMFLAG1 = False
'
' '**** TUMOUR MARKER REPORT
' '****ecg reports
'If selrrequisitnrs("MODALITYCD") = "PF" Then
'    If opt_patient.Value = True Then
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISPF.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'    '        If Trim(prepath) = "" Then
'    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'    '        Else
'    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'    '        End If
'        selta1012_qr(0) = selrrequisitnrs("CASENO")
'        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
'        CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
'        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
'            STR1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        Else
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'    Else
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISPFSTAFF.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'    '        If Trim(prepath) = "" Then
'    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'    '        Else
'    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'    '        End If
'        selta1012_qr(0) = selrrequisitnrs("CASENO")
'        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
'        If Not selta1012_rs.EOF Then
'            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & "" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
'        Else
'            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & "" & "'"
'        End If
'        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
'            STR1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        Else
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "PFT REPORT" & "'"
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'        CRYSTALREPORT1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'        CRYSTALREPORT1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
'        CRYSTALREPORT1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
'    End If
'    CRYSTALREPORT1.WindowState = crptMaximized
'    CRYSTALREPORT1.Action = 1
'End If
'If selrrequisitnrs("MODALITYCD") = "ST" Then
'    If opt_patient.Value = True Then
'       CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISST.RPT"
'       CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'    '        If Trim(prepath) = "" Then
'    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'    '        Else
'    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'    '        End If
'       CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
'       If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'           CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "STRESS TEST REPORT" & "'"
'           STR1 = ""
'           STR1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
'           If Len(Trim(selrrequisitnrs("enterby"))) > 0 Then
'               CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'           Else
'               CRYSTALREPORT1.Formulas(9) = "ENT = '" & "" & "'"
'           End If
'           CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'       Else
'           CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "STRESS TEST REPORT" & "'"
'           CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'       End If
'        CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'        SELSTREPORTQR(0) = req_chglen(VAR_REQNNO)
'        Set SELSTREPORTRS = SELSTREPORTQR.OpenResultset(1, 3)
'
'        If Len(Trim(SELSTREPORTRS("RISKFACTOR"))) > 0 Then
'            CRYSTALREPORT1.Formulas(13) = "rk = '" & "Risk Factor" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(13) = "rk = '" & "" & "'"
'        End If
'        If Len(Trim(SELSTREPORTRS("MAXHR"))) > 0 Then
'            CRYSTALREPORT1.Formulas(14) = "maxhr = '" & "Max HR" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(14) = "maxhr = '" & "" & "'"
'        End If
'        If Len(Trim(SELSTREPORTRS("MAXBP"))) > 0 Then
'            CRYSTALREPORT1.Formulas(15) = "maxBP = '" & "Max BP" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(15) = "maxBP = '" & "" & "'"
'        End If
'        If Len(Trim(SELSTREPORTRS("DOBPRODUCT"))) > 0 Then
'            CRYSTALREPORT1.Formulas(16) = "DP = '" & "Double Product" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(16) = "DP = '" & "" & "'"
'        End If
'        If Len(Trim(SELSTREPORTRS("THRACHIVED"))) > 0 Then
'            CRYSTALREPORT1.Formulas(17) = "thra = '" & "% of THR Achieved" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(17) = "thra = '" & "" & "'"
'        End If
'        If Len(Trim(SELSTREPORTRS("STTCHANGES"))) > 0 Then
'            CRYSTALREPORT1.Formulas(18) = "sttc = '" & "ST/T Changes" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(18) = "sttc = '" & "" & "'"
'        End If
'        If Len(Trim(SELSTREPORTRS("ARRYTHMIA"))) > 0 Then
'            CRYSTALREPORT1.Formulas(19) = "arrythmia = '" & "Arrythmia" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(19) = "arrythmia = '" & "" & "'"
'        End If
'        If Len(Trim(SELSTREPORTRS("METSACHVED"))) > 0 Then
'            CRYSTALREPORT1.Formulas(20) = "metsa = '" & "METS Achieved" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(20) = "metsa = '" & "" & "'"
'        End If
'        If Len(Trim(SELSTREPORTRS("IMPRESSION"))) > 0 Then
'            CRYSTALREPORT1.Formulas(21) = "rt = '" & "Reason for termination" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(21) = "rt = '" & "" & "'"
'        End If
'    Else
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISSTSTAFF.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'    '        If Trim(prepath) = "" Then
'    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
'    '        Else
'    '            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
'    '        End If
'        CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
'        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'           CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "STRESS TEST REPORT" & "'"
'           STR1 = ""
'           STR1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
'           If Len(Trim(selrrequisitnrs("enterby"))) > 0 Then
'               CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'           Else
'               CRYSTALREPORT1.Formulas(9) = "ENT = '" & "" & "'"
'           End If
'           CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        Else
'           CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "STRESS TEST REPORT" & "'"
'           CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'        SELSTREPORTQR(0) = req_chglen(VAR_REQNNO)
'        Set SELSTREPORTRS = SELSTREPORTQR.OpenResultset(1, 3)
'        If Not SELSTREPORTRS.EOF Then
'            If Len(Trim(SELSTREPORTRS("RISKFACTOR"))) > 0 Then
'                CRYSTALREPORT1.Formulas(13) = "rk = '" & "Risk Factor" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(13) = "rk = '" & "" & "'"
'            End If
'        Else
'            CRYSTALREPORT1.Formulas(13) = "rk = '" & "" & "'"
'        End If
'        If Not SELSTREPORTRS.EOF Then
'            If Len(Trim(SELSTREPORTRS("MAXHR"))) > 0 Then
'                CRYSTALREPORT1.Formulas(14) = "maxhr = '" & "Max HR" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(14) = "maxhr = '" & "" & "'"
'            End If
'        Else
'            CRYSTALREPORT1.Formulas(14) = "maxhr = '" & "" & "'"
'        End If
'        If Not SELSTREPORTRS.EOF Then
'            If Len(Trim(SELSTREPORTRS("MAXBP"))) > 0 Then
'                CRYSTALREPORT1.Formulas(15) = "maxBP = '" & "Max BP" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(15) = "maxBP = '" & "" & "'"
'            End If
'        Else
'            CRYSTALREPORT1.Formulas(15) = "maxBP = '" & "" & "'"
'        End If
'        If Not SELSTREPORTRS.EOF Then
'            If Len(Trim(SELSTREPORTRS("DOBPRODUCT"))) > 0 Then
'                CRYSTALREPORT1.Formulas(16) = "DP = '" & "Double Product" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(16) = "DP = '" & "" & "'"
'            End If
'        Else
'            CRYSTALREPORT1.Formulas(16) = "DP = '" & "" & "'"
'        End If
'        If Not SELSTREPORTRS.EOF Then
'            If Len(Trim(SELSTREPORTRS("THRACHIVED"))) > 0 Then
'                CRYSTALREPORT1.Formulas(17) = "thra = '" & "% of THR Achieved" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(17) = "thra = '" & "" & "'"
'            End If
'        Else
'            CRYSTALREPORT1.Formulas(17) = "thra = '" & "" & "'"
'        End If
'        If Not SELSTREPORTRS.EOF Then
'            If Len(Trim(SELSTREPORTRS("STTCHANGES"))) > 0 Then
'                CRYSTALREPORT1.Formulas(18) = "sttc = '" & "ST/T Changes" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(18) = "sttc = '" & "" & "'"
'            End If
'        Else
'            CRYSTALREPORT1.Formulas(18) = "sttc = '" & "" & "'"
'        End If
'        If Not SELSTREPORTRS.EOF Then
'            If Len(Trim(SELSTREPORTRS("ARRYTHMIA"))) > 0 Then
'                CRYSTALREPORT1.Formulas(19) = "arrythmia = '" & "Arrythmia" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(19) = "arrythmia = '" & "" & "'"
'            End If
'        Else
'            CRYSTALREPORT1.Formulas(19) = "arrythmia = '" & "" & "'"
'        End If
'        If Not SELSTREPORTRS.EOF Then
'            If Len(Trim(SELSTREPORTRS("METSACHVED"))) > 0 Then
'                CRYSTALREPORT1.Formulas(20) = "metsa = '" & "METS Achieved" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(20) = "metsa = '" & "" & "'"
'            End If
'        Else
'            CRYSTALREPORT1.Formulas(20) = "metsa = '" & "" & "'"
'        End If
'        If Not SELSTREPORTRS.EOF Then
'            If Len(Trim(SELSTREPORTRS("IMPRESSION"))) > 0 Then
'                CRYSTALREPORT1.Formulas(21) = "rt = '" & "Reason for termination" & "'"
'            Else
'                CRYSTALREPORT1.Formulas(21) = "rt = '" & "" & "'"
'            End If
'        Else
'            CRYSTALREPORT1.Formulas(21) = "rt = '" & "" & "'"
'        End If
'        CRYSTALREPORT1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'        CRYSTALREPORT1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
'        CRYSTALREPORT1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
'    End If
'    CRYSTALREPORT1.WindowState = crptMaximized
'    CRYSTALREPORT1.Action = 1
'End If
'If selrrequisitnrs("MODALITYCD") = "EK" Then
'    If opt_patient.Value = True Then
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISEC.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
''        If Trim(prepath) = "" Then
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
''        Else
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
''        End If
'        'CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
'        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "ECG REPORT" & "'"
'            STR1 = ""
'            STR1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
'            If Len(Trim(selrrequisitnrs("enterby"))) > 0 Then
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            Else
'                 CRYSTALREPORT1.Formulas(9) = "ENT = '" & "" & "'"
'            End If
'        Else
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "ECG REPORT" & "'"
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'        If SELECREPORTFLAG = False Then
'            SELECREPORTQR(0) = VAR_REQNNO
'            Set SELECREPORTRS = SELECREPORTQR.OpenResultset(1, 3)
'            SELECREPORTFLAG = True
'        Else
'            SELECREPORTQR(0) = VAR_REQNNO
'            SELECREPORTRS.MoreResults
'            SELECREPORTRS.Requery
'        End If
'        If Len(Trim(SELECREPORTRS("RHYTHM"))) > 3 Then
'            CRYSTALREPORT1.Formulas(13) = "rhythm = '" & "Rhythm" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(13) = "rhythm = '" & "" & "'"
'        End If
'        If (Trim(SELECREPORTRS("PWAVE"))) <> "" Then
'            CRYSTALREPORT1.Formulas(14) = "pwave = '" & "Pwave" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(14) = "pwave = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("TWAVE"))) > 3 Then
'            CRYSTALREPORT1.Formulas(15) = "twave = '" & "Twave" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(15) = "twave = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("STSEGMENT"))) > 3 Then
'            CRYSTALREPORT1.Formulas(16) = "stsegment = '" & "ST Segment" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(16) = "stsegment = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("RATE"))) > 0 Then
'            CRYSTALREPORT1.Formulas(17) = "rate = '" & "Rate" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(17) = "rate = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("ATRIALRATE"))) > 0 Then
'            CRYSTALREPORT1.Formulas(18) = "arate = '" & "Atrial Rate" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(18) = "arate = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("VENRATE"))) > 0 Then
'            CRYSTALREPORT1.Formulas(19) = "vrate = '" & "Ventricular Rate" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(19) = "vrate = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("PRINTERVAL"))) > 0 Then
'            CRYSTALREPORT1.Formulas(20) = "printerval = '" & "PR Interval" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(20) = "printerval = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("QRSMORPH"))) > 0 Then
'            CRYSTALREPORT1.Formulas(21) = "qrsm = '" & "QRS Morphology" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(21) = "qrsm = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("QRSDUR"))) > 0 Then
'            CRYSTALREPORT1.Formulas(22) = "qrd = '" & "QRS Duration" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(22) = "qrd = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("QTINTERVAL"))) > 0 Then
'            CRYSTALREPORT1.Formulas(23) = "qti = '" & "QT Interval" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(23) = "qti = '" & "" & "'"
'        End If
'    Else
'        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISEC.RPT"
'        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
''        If Trim(prepath) = "" Then
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
''        Else
''            CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
''        End If
'        'CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
'        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "ECG REPORT" & "'"
'            STR1 = ""
'            STR1 = "Entered By " & Trim(selrrequisitnrs("enterby"))
'            If Len(Trim(selrrequisitnrs("enterby"))) > 0 Then
'                CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'            Else
'                 CRYSTALREPORT1.Formulas(9) = "ENT = '" & "" & "'"
'            End If
'        Else
'            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "ECG REPORT" & "'"
'            CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'        End If
'        CRYSTALREPORT1.Formulas(12) = "USRID = '" & STR & "'"
'        If SELECREPORTFLAG = False Then
'            SELECREPORTQR(0) = VAR_REQNNO
'            Set SELECREPORTRS = SELECREPORTQR.OpenResultset(1, 3)
'            SELECREPORTFLAG = True
'        Else
'            SELECREPORTQR(0) = VAR_REQNNO
'            SELECREPORTRS.MoreResults
'            SELECREPORTRS.Requery
'        End If
'        If Len(Trim(SELECREPORTRS("RHYTHM"))) > 3 Then
'            CRYSTALREPORT1.Formulas(13) = "rhythm = '" & "Rhythm" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(13) = "rhythm = '" & "" & "'"
'        End If
'        If (Trim(SELECREPORTRS("PWAVE"))) <> "" Then
'            CRYSTALREPORT1.Formulas(14) = "pwave = '" & "Pwave" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(14) = "pwave = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("TWAVE"))) > 3 Then
'            CRYSTALREPORT1.Formulas(15) = "twave = '" & "Twave" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(15) = "twave = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("STSEGMENT"))) > 3 Then
'            CRYSTALREPORT1.Formulas(16) = "stsegment = '" & "ST Segment" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(16) = "stsegment = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("RATE"))) > 0 Then
'            CRYSTALREPORT1.Formulas(17) = "rate = '" & "Rate" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(17) = "rate = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("ATRIALRATE"))) > 0 Then
'            CRYSTALREPORT1.Formulas(18) = "arate = '" & "Atrial Rate" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(18) = "arate = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("VENRATE"))) > 0 Then
'            CRYSTALREPORT1.Formulas(19) = "vrate = '" & "Ventricular Rate" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(19) = "vrate = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("PRINTERVAL"))) > 0 Then
'            CRYSTALREPORT1.Formulas(20) = "printerval = '" & "PR Interval" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(20) = "printerval = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("QRSMORPH"))) > 0 Then
'            CRYSTALREPORT1.Formulas(21) = "qrsm = '" & "QRS Morphology" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(21) = "qrsm = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("QRSDUR"))) > 0 Then
'            CRYSTALREPORT1.Formulas(22) = "qrd = '" & "QRS Duration" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(22) = "qrd = '" & "" & "'"
'        End If
'        If Len(Trim(SELECREPORTRS("QTINTERVAL"))) > 0 Then
'            CRYSTALREPORT1.Formulas(23) = "qti = '" & "QT Interval" & "'"
'        Else
'            CRYSTALREPORT1.Formulas(23) = "qti = '" & "" & "'"
'        End If
'        CRYSTALREPORT1.Formulas(8) = "NAM = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'        CRYSTALREPORT1.Formulas(10) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
'        CRYSTALREPORT1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
'    End If
'    CRYSTALREPORT1.WindowState = crptMaximized
'    CRYSTALREPORT1.Action = 1
'End If
'
' '****ecg report
' '**** RIS REPORTS
'If selrrequisitnrs("MODALITYCD") = "CT" Or selrrequisitnrs("MODALITYCD") = "US" Or selrrequisitnrs("MODALITYCD") = "MM" Or selrrequisitnrs("MODALITYCD") = "MR" Or selrrequisitnrs("MODALITYCD") = "CR" Then
'    rq_grid_det(0) = Trim(VAR_REQNNO)
'    rs_grid_det.MoreResults
'    rs_grid_det.Requery
'    If rs_grid_det.RowCount > 0 And (Trim(rs_grid_det("REQNSTAT")) = "T" Or Trim(rs_grid_det("REQNSTAT")) = "F") Then
'        If opt_patient.Value = True Then
'            CRYSTALREPORT1.Formulas(0) = "provtyp= '" & Trim(rs_grid_det("ENTERBY")) & "'"
'            'cr_rep.Formulas(1) = "reporttext= '" & Trim(txt_report.Text) & "'"
'            'cr_rep.Formulas(0) = "impresstext= '" & Trim(txt_impression.Text) & "'"
'            'cr_rep.Formulas(1) = "finalrep= '" & Trim(txt_finalreport.Text) & "'"
'
'            SELRMODALITY_QR(0) = Mid(VAR_REQNNO, 4, 2)
'            SELRMODALITY_RS.MoreResults
'            SELRMODALITY_RS.Requery
'            If SELRMODALITY_RS.RowCount > 0 Then
'                CRYSTALREPORT1.Formulas(1) = "MODDES='" & SELRMODALITY_RS("MODLITYDES") & "'"
'            Else
'                CRYSTALREPORT1.Formulas(1) = "MODDES='" & "" & "'"
'            End If
'            CRYSTALREPORT1.Formulas(2) = "deptname= '" & "Department of Radiodiagnosis" & "'"
'            CRYSTALREPORT1.SelectionFormula = "{RISREPORTS.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'            'cr_rep.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'            ''cr_rep.Formulas(0) = "provtypist = '" & txt_type1.Text & "'"
'            STR1 = ip
'            STR2 = DATE11
'            STR3 = time1
'            STR1 = STR1 & Space(5) & STR2 & Space(5) & STR3
'            STR1 = usr_name & Space(3) & STR1
'            CRYSTALREPORT1.Formulas(3) = "us_id = '" & STR1 & "'"
'            CRYSTALREPORT1.Formulas(4) = "INOUT = '" & Trim(rs_grid_det("PATNSTAT")) & "'"
'            CRYSTALREPORT1.Formulas(5) = "reqno = '" & Left(msf_reqstat.Text, 16) & "'"
'            If Trim(rs_grid_det("REQNSTAT")) = "T" Or Trim(rs_grid_det("REQNSTAT")) = "F" Then
'                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\crpt_rep1.rpt"
'            Else
'                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\crpt_rep2.rpt"
'            End If
'        Else
'            CRYSTALREPORT1.Formulas(0) = "provtyp= '" & Trim(rs_grid_det("ENTERBY")) & "'"
'            'cr_rep.Formulas(1) = "reporttext= '" & Trim(txt_report.Text) & "'"
'            'cr_rep.Formulas(0) = "impresstext= '" & Trim(txt_impression.Text) & "'"
'            'cr_rep.Formulas(1) = "finalrep= '" & Trim(txt_finalreport.Text) & "'"
'
'            SELRMODALITY_QR(0) = Mid(VAR_REQNNO, 4, 2)
'            SELRMODALITY_RS.MoreResults
'            SELRMODALITY_RS.Requery
'            If SELRMODALITY_RS.RowCount > 0 Then
'                CRYSTALREPORT1.Formulas(1) = "MODDES='" & SELRMODALITY_RS("MODLITYDES") & "'"
'            Else
'                CRYSTALREPORT1.Formulas(1) = "MODDES='" & "" & "'"
'            End If
'            CRYSTALREPORT1.Formulas(2) = "deptname= '" & "Department of Radiodiagnosis" & "'"
'            CRYSTALREPORT1.SelectionFormula = "{RISREPORTS.REQNNO} = '" & VAR_REQNNO & "' AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'            'cr_rep.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"
'            ''cr_rep.Formulas(0) = "provtypist = '" & txt_type1.Text & "'"
'            STR1 = ip
'            STR2 = DATE11
'            STR3 = time1
'            STR1 = STR1 & Space(5) & STR2 & Space(5) & STR3
'            STR1 = usr_name & Space(3) & STR1
'            CRYSTALREPORT1.Formulas(3) = "us_id = '" & STR1 & "'"
'            CRYSTALREPORT1.Formulas(4) = "INOUT = '" & Trim(rs_grid_det("PATNSTAT")) & "'"
'            CRYSTALREPORT1.Formulas(5) = "reqno = '" & Left(msf_reqstat.Text, 16) & "'"
'            CRYSTALREPORT1.Formulas(6) = "PATNAME = '" & Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("FNAME")) & " " & Trim(SELEMPBASINFO_RS("MNAME")) & " " & Trim(SELEMPBASINFO_RS("LNAME")) & "'"
'            CRYSTALREPORT1.Formulas(7) = "STDOB = '" & SELEMPBASINFO_RS("GENDER") & "'"
'            CRYSTALREPORT1.Formulas(5) = "PATAGE = " & DateDiff("yyyy", SELEMPBASINFO_RS("DOB"), selrrequisitnrs("REQNDT"))
'            If Trim(rs_grid_det("REQNSTAT")) = "T" Or Trim(rs_grid_det("REQNSTAT")) = "F" Then
'                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\crpt_rep1STAFF.rpt"
'            Else
'                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\crpt_rep2STAFF.rpt"
'            End If
'        End If
'        CRYSTALREPORT1.WindowState = crptMaximized
'        CRYSTALREPORT1.Action = 1
'        Exit Sub
'    Else
'        Exit Sub
'    End If
'End If
' '**** RIS REPORTS
' '*************************************************************************
' '          wrote for microbiology ****************************************
' ' on 12/12/2002
' If selrrequisitnrs("MODALITYCD") = "BA" Or selrrequisitnrs("MODALITYCD") = "MB" Or selrrequisitnrs("MODALITYCD") = "XX" Or selrrequisitnrs("MODALITYCD") = "SE" Or selrrequisitnrs("MODALITYCD") = "CM" Then
'    Microbiology selrrequisitnrs("MODALITYCD")
' ElseIf selrrequisitnrs("MODALITYCD") = "2D" Then
'    Echocardiology selrrequisitnrs("MODALITYCD")
' ElseIf selrrequisitnrs("MODALITYCD") = "ED" Then
'    Endoscopy selrrequisitnrs("MODALITYCD")
' End If
' Exit Sub
'MSG:
'MsgBox Err.Description
'End Sub

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
msk_caseno.Text = UCase(msk_caseno.Text)
If opt_patient.Value = True Then
    cno_pad = padzero(5, "C", msk_caseno.Text, 3)
    msk_caseno.Text = cno_pad
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
        TXT_NAME.Visible = True
        LBL_NAME.Visible = True
        
        TXT_NAME.Text = Trim(SELTA101_RS("TAG")) & " " & Trim(SELTA101_RS("fIRST_NAME")) & " " & Trim(SELTA101_RS("MID_NAME")) & " " & Trim(SELTA101_RS("LAST_NAME"))
    End If
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
    cno_pad = msk_caseno.FormattedText
    If SELRREQUISITNFLAG1 = False Then
        SELRREQUISITN_QR(0) = msk_caseno.FormattedText
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITNFLAG1 = True
    Else
        SELRREQUISITN_QR(0) = msk_caseno.FormattedText
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
   msk_caseno.Enabled = False
   cbo_req.Enabled = False
   cbo_modality.Enabled = True
   
End Sub

Private Sub Opt_case_Click()

   Opt_all.Enabled = False
   MaskEd_from_date.Enabled = False
    cbo_modality.Enabled = False
   MaskEd_to_date.Enabled = False
   msk_caseno.Enabled = True
   cbo_req.Enabled = True
   
    msk_caseno.SetFocus
   
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
    msk_caseno.Text = ""
    cbo_req.Text = ""
    cbo_modality.ListIndex = -1
'    msf_reqstat.clear
    
    Opt_all.Enabled = False
    Opt_case.Enabled = False
    msk_caseno.Enabled = False
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


Private Sub Opt_patient_Click()
If opt_patient.Value = True Then
     'lbl_caseno.Caption = "Case No."
     msk_caseno.Mask = "??>/#####"
    'LBL_DEPENDENT.Visible = False
    'cmb_dependents.Visible = False
     msk_caseno.Enabled = True
     'msk_caseno.SetFocus
   
End If
End Sub

Private Sub opt_staff_Click()
On Error GoTo MSG:
If opt_staff.Value = True Then
    Opt_case.Caption = "CC NO."
    'lbl_caseno.Caption = "  CC No."
    msk_caseno.Mask = "########"
    'LBL_DEPENDENT.Visible = True
    'cmb_dependents.Visible = True
    msk_caseno.Enabled = True
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
                                
                                SELDREQNOQR1(0) = RS_ALLSP("REQNNO")
                                Set SELDREQNORS1 = SELDREQNOQR1.OpenResultset(1, 3)
                                If SELDREQNORS1.RowCount > 0 Then
                                    msf_reqstat.CellBackColor = vbGreen
                                    msf_reqstat.TextMatrix(srno, 8) = req_chglen(SELDREQNORS1("REQNO"))
                                End If
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


