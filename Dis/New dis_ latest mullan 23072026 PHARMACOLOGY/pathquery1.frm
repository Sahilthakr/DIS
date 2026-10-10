VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form h 
   Caption         =   "Pathology Queries"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12195
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Order by"
      Height          =   615
      Left            =   7680
      TabIndex        =   26
      Top             =   1920
      Visible         =   0   'False
      Width           =   3855
      Begin VB.OptionButton OPT_NMAT 
         Caption         =   "N.Material"
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OPT_REQNO 
         Caption         =   "Requisition No."
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.ComboBox cmb_dept 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5040
      TabIndex        =   20
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ComboBox Cmb_modality 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8640
      TabIndex        =   19
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   3855
      Begin VB.ComboBox cmb_docname 
         Height          =   315
         Left            =   1440
         TabIndex        =   24
         Top             =   2040
         Width           =   2295
      End
      Begin VB.OptionButton OPT_DOCTORWISE 
         Caption         =   "Doctor Wise "
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
      End
      Begin VB.OptionButton OPT_RFCASE 
         Caption         =   "RF CASES"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton OPT_NIP 
         Caption         =   "NIP "
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton opt_pr 
         Caption         =   "Print the 'SP' Reports between two dates"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   3135
      End
      Begin VB.OptionButton opt_pd 
         Caption         =   "Patient Details between two dates"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lbl_doc 
         AutoSize        =   -1  'True
         Caption         =   "Doctor Name"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   2160
         Width           =   1125
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   7680
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
         MouseIcon       =   "pathquery1.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   0
         Top             =   120
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
         MouseIcon       =   "pathquery1.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   120
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
         MouseIcon       =   "pathquery1.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   120
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
         MouseIcon       =   "pathquery1.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf_reqstat 
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "CLICK ON THE REQUISTION NO."
      Top             =   3360
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   10
      BackColor       =   16777215
      ForeColor       =   0
      ForeColorFixed  =   1194390
      WordWrap        =   -1  'True
      Enabled         =   -1  'True
      AllowUserResizing=   1
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
   Begin MSMask.MaskEdBox MaskEd_to_date 
      Height          =   375
      Left            =   9270
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
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
   Begin MSMask.MaskEdBox MaskEd_from_date 
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   2760
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   0
      PromptInclude   =   0   'False
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
      Left            =   960
      TabIndex        =   29
      Top             =   120
      Width           =   10455
   End
   Begin VB.Label lbl_dept 
      AutoSize        =   -1  'True
      Caption         =   "Department"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3960
      TabIndex        =   22
      Top             =   1440
      Width           =   990
   End
   Begin VB.Label lbl_mod 
      AutoSize        =   -1  'True
      Caption         =   "Modality"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7800
      TabIndex        =   21
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label lbl_br 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   7320
      Width           =   4575
   End
   Begin VB.Label lbl_drop 
      Height          =   495
      Left            =   4560
      TabIndex        =   17
      Top             =   7200
      Width           =   4575
   End
   Begin VB.Label lbl_lr 
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   6720
      Width           =   4455
   End
   Begin VB.Label LBL_RPTNO 
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Width           =   4335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   3960
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
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
      Left            =   8280
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
End
Attribute VB_Name = "h"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RQ_ALLREQ1 As rdoQuery
Dim rs_allreq As rdoResultset
Dim rs_allreqflag As Boolean

Dim RQ_ALLREQ2 As rdoQuery
Dim rs_allreq2 As rdoResultset
Dim rs_allreqflag2 As Boolean


Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset
Dim selta101rsflag As Boolean

Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim seldrequisitnrsflag As Boolean

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

Dim selrrequisitnqr As rdoQuery
Dim selrrequisitnrs As rdoResultset
Dim SELRREQUISITNRSFLAG As Boolean

Dim SELRREQUISITN_QR As rdoQuery
Dim SELRREQUISITN_RS As rdoResultset
Dim SELRREQUISITN_RSFLAG As Boolean

Dim SELFRREQUISITN_QR As rdoQuery
Dim SELFRREQUISITN_RS As rdoResultset
Dim SELFRREQUISITN_RSFLAG As Boolean

Dim SELMA106_QR As rdoQuery
Dim SELMA106_RS As rdoResultset
Dim SELMA106FLAG As Boolean

Dim SELRMODALITY_QR As rdoQuery
Dim SELRMODALITY_RS As rdoResultset
Dim SELRMODALITYFLAG As Boolean

Dim seldissecfin_qr As rdoQuery
Dim seldissecfin_rs As rdoResultset
Dim seldissecfinrsflag As Boolean

Dim seldimpresion_qr As rdoQuery
Dim seldimpresion_rs As rdoResultset
Dim SELDIMPRESIONFLAG As Boolean
Dim dept As String
Dim moda As String

Dim vpp As Integer
Dim lr As Integer
Dim bar As Integer
Dim drop1 As Integer


''NEW
Dim SELMA106_QR1 As rdoQuery
Dim SELMA106_RS1 As rdoResultset
Dim SELMA106_FLAG1 As Boolean

Dim SELRMODALITY_QR1 As rdoQuery
Dim SELRMODALITY_RS1 As rdoResultset
Dim SELRMODALITY_FLAG1 As Boolean

Dim SELDSPLF_QR As rdoQuery
Dim SELDSPLF_RS As rdoResultset
Dim SELDSPLF_FLAG As Boolean

''FOR STAFF
Dim SELEMPBASINFO_QR As rdoQuery
Dim SELEMPBASINFO_RS As rdoResultset
Dim SELEMPBASINFO_FLAG As Boolean

Private Sub cmb_dept_Click()
Dim i As Integer
i = InStr(cmb_dept.Text, "$")
dept = Mid(cmb_dept.Text, i + 1, 4)
If SELRMODALITYFLAG = False Then
    SELRMODALITY_QR(0) = Trim(dept)
    Set SELRMODALITY_RS = SELRMODALITY_QR.OpenResultset(1, 3)
    SELRMODALITYFLAG = True
Else
    SELRMODALITY_QR(0) = Trim(dept)
    SELRMODALITY_RS.MoreResults
    SELRMODALITY_RS.Requery
End If
If SELRMODALITY_RS.RowCount > 0 Then
    cmb_modality.clear
    While Not SELRMODALITY_RS.EOF
        cmb_modality.AddItem Trim(SELRMODALITY_RS("modlitydes")) & "$" & Trim(SELRMODALITY_RS("modalitycd"))
        SELRMODALITY_RS.MoveNext
    Wend
End If
cmb_modality.SetFocus
End Sub



Private Sub cmb_modality_Click()
'SR.NO. |^ REQUISITION NO./ DATE      |^  PATH NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       N.MATERIAL         |^       DIAGNOSIS       |^     FILE NO.
i = InStr(cmb_modality.Text, "$")
moda = Mid(cmb_modality.Text, i + 1, 2)
If Trim(moda) = "CN" Or Trim(moda) = "CG" Then
Frame2.Visible = True
End If
   msf_reqstat.clear
 If opt_pd.Value = True Then
    msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  PATH NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       N.MATERIAL         |^       DIAGNOSIS       |^     FILE NO.   |^ Report  |^ Dr. Name "
 End If
 If opt_pr.Value = True Then
    msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  PATH NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       N.MATERIAL         |^       DIAGNOSIS       |^     FILE NO.   |^ Report  |^ Dr. Name "
 End If
 If OPT_NIP.Value = True Then
    If moda = "SP" Or moda = "FA" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  PATH NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       REF BY DOC         |^       IMPRESSION       |^     NIP     |^  FINALISED BY    |^ RECEIPT NO.  "
    End If
    If moda = "CG" Or moda = "CN" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  CYTO NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       REF BY DOC         |^       REF TO DOC      |^     NIP    |^  FINALISED BY    |^ RECEIPT NO.  "
    End If
    If moda = "BA" Or moda = "MB" Or moda = "SE" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  MIC NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       REF BY DOC         |^       REF TO DOC       |^     NIP     |^  FINALISED BY    |^ RECEIPT NO.  "
    End If
    If moda = "BC" Or moda = "EL" Or moda = "TM" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  LAB NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       REF BY DOC         |^       REF TO DOC       |^     NIP     |^  FINALISED BY    |^ RECEIPT NO.  "
    End If
 End If
 If OPT_RFCASE.Value = True Then
    If moda = "SP" Or moda = "IH" Or moda = "FA" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  PATH NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       N.MATERIAL         |^       DIAGNOSIS       |^     NIP"
    End If
    If moda = "CG" Or moda = "CN" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  CYTO NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       N.MATERIAL         |^       DIAGNOSIS       |^     NIP"
    End If
    If moda = "BA" Or moda = "MB" Or moda = "SE" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  MIC NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       N.MATERIAL         |^       DIAGNOSIS       |^     NIP"
    End If
    
    If moda = "BC" Or moda = "EL" Or moda = "TM" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  LAB NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       N.MATERIAL         |^       DIAGNOSIS       |^     NIP"
    End If
    
 End If

If opt_pd.Value = True Then
    If moda = "SP" Or moda = "IH" Or moda = "FA" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  PATH NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       N.MATERIAL         |^       DIAGNOSIS       |^     FILE NO."
    End If
    If moda = "CG" Or moda = "CN" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  CYTO NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       N.MATERIAL         |^       DIAGNOSIS       |^    FILE NO."
    End If
    If moda = "BA" Or moda = "MB" Or moda = "SE" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  MIC NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       N.MATERIAL         |^       DIAGNOSIS       |^     FILE NO."
    End If
    
    If moda = "BC" Or moda = "EL" Or moda = "TM" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  LAB NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       N.MATERIAL         |^       DIAGNOSIS       |^     FILE NO."
    End If
    
 End If

If OPT_DOCTORWISE.Value = True Then
    If moda = "SP" Or moda = "IH" Or moda = "FA" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  PATH NO.          |^ CASE NO.    |^                   NAME                                |^   SEX    |^   AGE    |^     NIP       |^       RECEIPT NO.        "
    End If
    If moda = "CG" Or moda = "CN" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  CYTO NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^       NIP     |^       RECEIPT NO.       "
    End If
    If moda = "BA" Or moda = "MB" Or moda = "SE" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  MIC NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^        NIP     |^        RECEIPT NO.       "
    End If
    
    If moda = "BC" Or moda = "EL" Or moda = "TM" Then
        msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE      |^  LAB NO.          |^ CASE NO.    |^                    NAME                                |^   SEX    |^   AGE    |^        NIP     |^      RECEIPT NO.       "
    End If
    
 End If


MaskEd_from_date.Enabled = True
MaskEd_from_date.SetFocus
End Sub

Private Sub cmd_cancel_Click()
    Frame2.Visible = False
    cmb_dept.Text = ""
    cmb_modality.Text = ""
    MaskEd_from_date.Mask = ""
    MaskEd_from_date.Text = ""
    MaskEd_from_date.Mask = "##/##/####"
    
    MaskEd_to_date.Mask = ""
    MaskEd_to_date.Text = ""
    MaskEd_to_date.Mask = "##/##/####"
    MaskEd_from_date.Enabled = False
    MaskEd_to_date.Enabled = False
    cmd_cancel.Enabled = False
    cmd_print.Enabled = False
    cmd_view.Enabled = True
    cmd_exit.Enabled = True
    opt_pd.Value = False
    opt_pr.Value = False
    msf_reqstat.Rows = 1
    LBL_RPTNO.Caption = ""
    lbl_lr.Caption = ""
    lbl_br.Caption = ""
    lbl_drop.Caption = ""
    OPT_NIP.Value = False
    OPT_RFCASE.Value = False
    OPT_DOCTORWISE.Value = False
    cmb_docname.clear
    msf_reqstat.clear
End Sub

Private Sub cmd_exit_Click()
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_print_Click()
Dim str1 As String
Dim str2 As String
Dim STR3 As String
 
On Error GoTo MSG:

DoEvents
Dim STR5 As String

msf_reqstat.row = 1
msf_reqstat.col = 1
Dim VAR_REQNNO As String
Dim var_yr_f As Integer, var_yr_t As Integer
Dim var_mon_f As Integer, var_mon_t As Integer
Dim var_day_f As Integer, var_day_t As Integer
Dim var_date_f As String, var_date_t As String
STR5 = "F"
If opt_pd.Value = True Then
    var_date_f = Year(MaskEd_from_date.FormattedText) & "," & month(MaskEd_from_date.FormattedText) & "," & Day(MaskEd_from_date.FormattedText)
    var_date_t = Year(MaskEd_to_date.FormattedText) & "," & month(MaskEd_to_date.FormattedText) & "," & Day(MaskEd_to_date.FormattedText)
    CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= '" & Trim(moda) & " '" ' AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
    'If Trim(moda) = "SP" Then
        'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") "   'AND {RREQUISITN.MODALITYCD}= '" & Trim(cbo_modality) & " ' AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
    CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\PATHreg.rpt"
    'End If
    If OPT_REQNO.Value = True Then
        If Trim(moda) = "CN" Then
            'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= 'CN'" ' ORDER BY REQNO" 'AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\MISCN.rpt"
        End If
        If Trim(moda) = "CG" Then
            'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= 'CG'" 'AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\MISCN.rpt"
        End If
    End If
    If OPT_NMAT.Value = True Then
        If Trim(moda) = "CN" Then
            'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= 'CN'" ' ORDER BY REQNO" 'AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\MISCN.rpt"
        End If
        If Trim(moda) = "CG" Then
            'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNDT} >= Date(" & var_date_f & ") AND {RREQUISITN.REQNDT} <= Date(" & var_date_t & ") AND {RREQUISITN.MODALITYCD}= 'CG'" 'AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\MISCN.rpt"
        End If
    End If
        str1 = ip
        str2 = DATE11
        STR3 = time1
        str1 = str1 & Space(5) & str2 & Space(5) & STR3
        
        'CrystalReport1.SelectionFormula = "{RREQUISITN.PREPORTDT} = Date(" & var_date_f & ") AND {RREQUISITN.MODALITYCD}= 'SP'" ' AND {RREQUISITN.REQNSTAT} <>  '" & STR5 & "' "
        CRYSTALREPORT1.Formulas(0) = "fromdt = '" & MaskEd_from_date.FormattedText & "'"
        CRYSTALREPORT1.Formulas(1) = "todt = '" & MaskEd_to_date.FormattedText & "'"
        CRYSTALREPORT1.WindowState = crptMaximized
        CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CRYSTALREPORT1.Action = 1
End If
'****
If rs_allreqflag2 = False Then
     RQ_ALLREQ2(0) = GET_REVDATE(MaskEd_from_date.FormattedText)
     RQ_ALLREQ2(1) = GET_REVDATE(MaskEd_to_date.FormattedText)
     RQ_ALLREQ2(2) = moda
     Set rs_allreq2 = RQ_ALLREQ2.OpenResultset(2, 3)
     rs_allreqflag2 = True
    
 Else
     RQ_ALLREQ2(0) = GET_REVDATE(MaskEd_from_date.FormattedText)
     RQ_ALLREQ2(1) = GET_REVDATE(MaskEd_to_date.FormattedText)
     RQ_ALLREQ2(2) = moda
     rs_allreq2.MoreResults
     rs_allreq2.Requery
End If
If opt_pr.Value = True Then
    If SELDSPLF_FLAG = False Then
         SELDSPLF_QR(0) = GET_REVDATE(MaskEd_from_date.FormattedText)
         SELDSPLF_QR(1) = GET_REVDATE(MaskEd_to_date.FormattedText)
         SELDSPLF_QR(2) = moda
         Set SELDSPLF_RS = SELDSPLF_QR.OpenResultset(2, 3)
         SELDSPLF_FLAG = True
        
     Else
         SELDSPLF_QR(0) = GET_REVDATE(MaskEd_from_date.FormattedText)
         SELDSPLF_QR(1) = GET_REVDATE(MaskEd_to_date.FormattedText)
         SELDSPLF_QR(2) = moda
         SELDSPLF_RS.MoreResults
         SELDSPLF_RS.Requery
    End If
''...*****TEST
    Dim XX As String
    XX = ""
    'XX = "{DSPLF.REQNNO}='"
    If SELDSPLF_RS.RowCount > 0 Then
        SELDSPLF_RS.MoveFirst
        LBL_RPTNO.Caption = "No. of Reports found Between two dates is " & SELDSPLF_RS.RowCount
        While Not SELDSPLF_RS.EOF
                'CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} IN( '" & rs_allreq2("reqnno") & "',"
                XX = XX & "{DSPLF.REQNNO}='" & SELDSPLF_RS("reqnno") & "' OR "
          SELDSPLF_RS.MoveNext
        Wend
    End If
            '*****
        XX = Mid(XX, 1, Len(XX) - 3)
        CRYSTALREPORT1.SelectionFormula = XX
        CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENMISTEST.RPT"
        'rs_allreq2.MoveFirst
'        If Not rs_allreq2.EOF Then
'        Do While Not rs_allreq2.EOF
                'If rs_allreq2("MODALITYCD") = "SP" Then
                     ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
                            'If rs_allreq2("COMMITFLAG") = "Y" Then
                             ' SELDADDENDUMQR(0) = rs_allreq2("reqnno")
                             '' Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
                               'If SELDADDENDUMRS.RowCount > 0 Then
                                '   CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
                              ' Else
                               '    CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
                               'End If
                          ' Else
                           '    CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENMISTEST.RPT"
                           'End If
                           'CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(rs_allreq2("reqnno")) & "'"
                           'seldrequisitn_qr(0) = rs_allreq2("reqnno")
                           'Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'                           If seldrequisitn_rs.RowCount > 0 Then
'                                CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
'                           Else
'                               CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'                           End If
                           'selta1012_qr(0) = rs_allreq2("CASENO")
                           'Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
                           'If selta1012_rs.RowCount > 0 Then
                            '   CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
                           'Else
                            '   CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & "" & "'"
                           'End If
                           'If seldrequisitnrsflag = False Then
                            '   seldrequisitn_qr(0) = rs_allreq2("reqnno")
                             '  Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                              ' seldrequisitnrsflag = True
                            'Else
                             '  seldrequisitn_qr(0) = rs_allreq2("reqnno")
                              ' seldrequisitn_rs.MoreResults
                               'seldrequisitn_rs.Requery
                            'End If

                           'SELMA301_QR(0) = seldrequisitn_rs("sITE")
                           'Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
                           'If SELMA301_RS.RowCount > 0 Then
                            '   CRYSTALREPORT1.Formulas(3) = "SIT = '" & Trim(SELMA301_RS("SITEDESC")) & "'"
                           'Else
                            '   CRYSTALREPORT1.Formulas(3) = "SIT = '" & "" & "'"
                           'End If
'                           If seldrequisitn_rs.RowCount > 0 Then
'                               CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & Trim(seldrequisitn_rs("OPPROCEDUR")) & "'"
'                           Else
'                               CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & "" & "'"
'                           End If
'                           STR1 = ""
                           'If Trim(rs_allreq2("REQNSTAT")) = "T" Then
                               'CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
                            '   STR1 = "Entered By " & Trim(rs_allreq2("enterby"))
                             '  CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
                           'Else
                               'CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
                            '   CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
                           'End If
                           'SELDFREPORT_QR(0) = rs_allreq2("reqnno")
                           'Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
                           'If SELDFREPORT_RS.RowCount > 0 Then
                            '   If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
                             '      CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
                              ' Else
                               '    CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "" & "'"
                               'End If
'                               If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
'                                   CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
'                               Else
'                                   CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "" & "'"
'                               End If
'                               If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
'                                   CRYSTALREPORT1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
'                               Else
'                                   CRYSTALREPORT1.Formulas(10) = "GROSDESC = '" & "" & "'"
'                               End If
'                               If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
'                                   CRYSTALREPORT1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
'                               Else
'                                   CRYSTALREPORT1.Formulas(11) = "MICRODESC = '" & "" & "'"
'                               End If
'                               If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
'                                   CRYSTALREPORT1.Formulas(12) = "IMPN = '" & "Impression" & "'"
'                               Else
'                                   CRYSTALREPORT1.Formulas(12) = "IMPN = '" & "" & "'"
'                               End If
'
'                           Else
'                               CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "" & "'"
'                           End If
                           'STR = usr_name & " " & ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                           'CRYSTALREPORT1.Formulas(8) = "USRID = '" & STR & "'"

                           'CrystalReport1.Formulas(10) = "GROSDESC = '" & "Gross description" & "'"
                           'CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
                           'CrystalReport1.Destination = crptToPrinter
                    ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
                    'End If
'
            'rs_allreq2.MoveNext
'        Loop
'    End If
            CRYSTALREPORT1.WindowState = crptMaximized
    '        CRYSTALREPORT1.WindowState = crptMaximized
            CRYSTALREPORT1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

            CRYSTALREPORT1.Action = 1

            '*****
'            rs_allreq2.MoveNext
'        Wend
    'End If
End If




'    If rs_allreq2.RowCount > 0 Then
'        LBL_RPTNO.Caption = "No. of Reports found Between two dates is " & rs_allreq2.RowCount
'        While Not rs_allreq2.EOF
'                CRYSTALREPORT1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & rs_allreq2("reqnno") & "'"
'            '*****
'                If rs_allreq2("MODALITYCD") = "SP" Then
'                     ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
'                            If rs_allreq2("COMMITFLAG") = "Y" Then
'                              SELDADDENDUMQR(0) = rs_allreq2("reqnno")
'                              Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
'                               If SELDADDENDUMRS.RowCount > 0 Then
'                                   CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
'                               Else
'                                   CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
'                               End If
'                           Else
'                               CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\FROZENMIS.RPT"
'                           End If
'                           CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(rs_allreq2("reqnno")) & "'"
'                           seldrequisitn_qr(0) = rs_allreq2("reqnno")
'                           Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'                           If seldrequisitn_rs.RowCount > 0 Then
'                                CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
'                           Else
'                               CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
'                           End If
'                           selta1012_qr(0) = rs_allreq2("CASENO")
'                           Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
'                           If selta1012_rs.RowCount > 0 Then
'                               CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
'                           Else
'                               CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & "" & "'"
'                           End If
'                           If seldrequisitnrsflag = False Then
'                               seldrequisitn_qr(0) = rs_allreq2("reqnno")
'                               Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
'                               seldrequisitnrsflag = True
'                            Else
'                               seldrequisitn_qr(0) = rs_allreq2("reqnno")
'                               seldrequisitn_rs.MoreResults
'                               seldrequisitn_rs.Requery
'                            End If
'
'                           SELMA301_QR(0) = seldrequisitn_rs("sITE")
'                           Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
'                           If SELMA301_RS.RowCount > 0 Then
'                               CRYSTALREPORT1.Formulas(3) = "SIT = '" & Trim(SELMA301_RS("SITEDESC")) & "'"
'                           Else
'                               CRYSTALREPORT1.Formulas(3) = "SIT = '" & "" & "'"
'                           End If
'                           If seldrequisitn_rs.RowCount > 0 Then
'                               CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & Trim(seldrequisitn_rs("OPPROCEDUR")) & "'"
'                           Else
'                               CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & "" & "'"
'                           End If
'                           STR1 = ""
'                           If Trim(rs_allreq2("REQNSTAT")) = "T" Then
'                               CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
'                               STR1 = "Entered By " & Trim(rs_allreq2("enterby"))
'                               CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'                           Else
'                               CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "HISTOPATHOLOGY REPORT" & "'"
'                               CRYSTALREPORT1.Formulas(9) = "ENT = '" & STR1 & "'"
'                           End If
'                           SELDFREPORT_QR(0) = rs_allreq2("reqnno")
'                           Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
'                           If SELDFREPORT_RS.RowCount > 0 Then
'                               If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
'                                   CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "Frozen Section" & "'"
'                               Else
'                                   CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "" & "'"
'                               End If
'                               If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
'                                   CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
'                               Else
'                                   CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "" & "'"
'                               End If
'                               If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
'                                   CRYSTALREPORT1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
'                               Else
'                                   CRYSTALREPORT1.Formulas(10) = "GROSDESC = '" & "" & "'"
'                               End If
'                               If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
'                                   CRYSTALREPORT1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
'                               Else
'                                   CRYSTALREPORT1.Formulas(11) = "MICRODESC = '" & "" & "'"
'                               End If
'                               If Len(Trim(SELDFREPORT_RS("Impression"))) > 2 Then
'                                   CRYSTALREPORT1.Formulas(12) = "IMPN = '" & "Impression" & "'"
'                               Else
'                                   CRYSTALREPORT1.Formulas(12) = "IMPN = '" & "" & "'"
'                               End If
'
'                           Else
'                               CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "" & "'"
'                           End If
'                           STR = usr_name & " " & ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
'                           CRYSTALREPORT1.Formulas(8) = "USRID = '" & STR & "'"
'
'                           'CrystalReport1.Formulas(10) = "GROSDESC = '" & "Gross description" & "'"
'                           'CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
'                           'CrystalReport1.Destination = crptToPrinter
'                    ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
'                    End If
'
'            CRYSTALREPORT1.Action = 1
'
'            '*****
'            rs_allreq2.MoveNext
'        Wend
'    End If
'End If


'****
Exit Sub
MSG:
MsgBox Err.Description

End Sub

Private Sub cmd_view_Click()
MaskEd_from_date.Enabled = True
MaskEd_to_date.Enabled = True
cmd_cancel.Enabled = True
cmd_print.Enabled = True
cmd_view.Enabled = False
cmd_exit.Enabled = True
Frame2.Visible = False
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Form_Load()
If PATHQUERIESFLAG = False Then
PATHQUERIESFLAG = True
datetime

'''NEW
str = "SELECT * FROM PATOOLSLIB.MA106 WHERE DEPT_CODE='NZZ'"
Set SELMA106_QR1 = rdb.CreateQuery("SELMA106_STR111", str)

str = "SELECT * FROM RISLIB.RMODALITY WHERE MODALITYCD='SP'"
Set SELRMODALITY_QR1 = rdb.CreateQuery("SELRMODALITY_STR111", str)


str = "select * from Rislib.RREQUISITN where  REQNDT between ? and  ?  AND MODALITYCD=? AND (REQNSTAT='T' OR REQNSTAT='F') order by caseno"
Set RQ_ALLREQ2 = rdb.CreateQuery("allreq2PQ2", str)

str = "select * from Dislib.DSPLF where  REQNDT between ? and  ?  AND MODALITYCD=? AND (REQNSTAT='T' OR REQNSTAT='F') order by reqndt desc"
Set SELDSPLF_QR = rdb.CreateQuery("SELDSPLF_STR", str)

str = "select * from patoolslib.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ')"
Set SELMA106_QR = rdb.CreateQuery("selma106_qr11", str)

str = "select * from rislib.rmodality where dept_code=?"
Set SELRMODALITY_QR = rdb.CreateQuery("selrmodality_qr1", str)

'Set RQ_ALLREQ1 = rdb.CreateQuery("allreq2PQ1", "select * from Rislib.RREQUISITN where (reqnstat ='T'  or reqnstat='F') AND REQNDT between ? and  ?  AND MODALITYCD=? order by reqndt desc")
Set RQ_ALLREQ1 = rdb.CreateQuery("allreq2PQ1", "select * from Rislib.RREQUISITN where  REQNDT between ? and  ?  AND MODALITYCD=? order by reqndt desc")


str = "SELECT * FROM RISLIB.RREQUISITN  WHERE   MODALITYCD=? AND  (RECORDNO=1 OR RECORDNO=2 OR RECORDNO=3 OR RECORDNO=4) AND REQNDT BETWEEN ? AND ? "
Set selrrequisitnqr = rdb.CreateQuery("strss", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE CASENO LIKE 'RF%' and  REQNDT BETWEEN ? AND ? and MODALITYCD=?"
Set SELRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN_QR2", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? and MODALITYCD=? AND FREPORTBY=?"
Set SELFRREQUISITN_QR = rdb.CreateQuery("SELRREQUISITN_QRF2", str)

str = "SELECT * FROM PATOOLSLIB.TA101 WHERE CAsE_NO=?"
Set SELTA101_QR = rdb.CreateQuery("RSELTA101_QRPQ1", str)

''FOR STAFF
str = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFO_STR", str)


str = "select * from dislib.drequisitn where reqno=?"
Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn1_qr1PQ1", str)


str = "select * from patoolslib.ta1012 where case_no=?"
Set selta1012_qr = rdb.CreateQuery("1selta10121_qrpq1", str)

str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
Set SELMA301_QR = rdb.CreateQuery("SELMA3011_QRpq1", str)

str = "SELECT * FROM DISLIB.DFREPORT WHERE REQNO=?"
Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT1_QRpq1", str)

str = "SELECT * FROM DISLIB.DADDENDUM WHERE REQNO=?"
Set SELDADDENDUMQR = rdb.CreateQuery("SELDADDENDUMQRpq1", str)

str = "SELECT * FROM DISLIB.DREQUISITN WHERE SPIHREQNO=?"
Set SELDREQNOQR1 = rdb.CreateQuery("S1ELDREQNOQR11", str)

str = "select * from dislib.dissec where FINALby='YES' AND MODALITY='SP'"
Set seldissecfin_qr = rdb.CreateQuery("seldissecFIN_qr21", str)

str = "select * from dislib.dfreport where reqno=?"
Set seldimpresion_qr = rdb.CreateQuery("seldimpresion_qr121", str)

End If
If SELMA106FLAG = False Then
    Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)
    SELMA106FLAG = True
Else
    SELMA106_RS.MoreResults
    SELMA106_RS.Requery
End If
If SELMA106_RS.RowCount > 0 Then
    cmb_dept.clear
    While Not SELMA106_RS.EOF
        cmb_dept.AddItem Trim(SELMA106_RS("dept_name")) & "$" & Trim(SELMA106_RS("dept_code"))
        SELMA106_RS.MoveNext
    Wend
End If
MaskEd_from_date.Enabled = False
MaskEd_to_date.Enabled = False
cmd_cancel.Enabled = False
cmd_print.Enabled = False
cmd_view.Enabled = True
cmd_exit.Enabled = True
opt_pd.Value = False
opt_pr.Value = False
If SELMA106_FLAG1 = False Then
    Set SELMA106_RS1 = SELMA106_QR1.OpenResultset(1, 3)
    SELMA106_FLAG1 = False
Else
    SELMA106_RS1.MoreResults
    SELMA106_RS1.Requery
End If
If SELRMODALITY_FLAG1 = False Then
    Set SELRMODALITY_RS1 = SELRMODALITY_QR1.OpenResultset(1, 3)
    SELRMODALITY_FLAG1 = False
Else
    SELRMODALITY_RS1.MoreResults
    SELRMODALITY_RS1.Requery
End If
End Sub

Private Sub MaskEd_from_date_LostFocus()
If InStr(MaskEd_from_date.FormattedText, "_") > 0 Then
            MaskEd_from_date.Mask = ""
            MaskEd_from_date.Text = ""
            MaskEd_from_date.Mask = "##/##/####"
            Exit Sub
End If

If IsDate(GET_REVDATE(MaskEd_from_date.FormattedText)) = False Then
             MsgBox "Please ENTER the correct DATE."
             MaskEd_from_date.Mask = ""
             MaskEd_from_date.Text = ""
             MaskEd_from_date.Mask = "##/##/####"
'             MaskEd_from_date.Text = ""
             MaskEd_from_date.SetFocus
             'Exit Sub
             Else
             MaskEd_to_date.Enabled = True
             'MaskEd_to_date.SetFocus
End If
End Sub

Private Sub MaskEd_to_date_LostFocus()
Dim TEMP As String
Dim YES As Integer
Dim NO As Integer
YES = 0
NO = 0
On Error GoTo MSG:
If InStr(MaskEd_to_date.FormattedText, "_") > 0 Then
    MaskEd_to_date.Mask = ""
    MaskEd_to_date.Text = ""
    MaskEd_to_date.Mask = "##/##/####"
    Exit Sub
End If
If IsDate(GET_REVDATE(MaskEd_to_date.FormattedText)) = False Then
    MaskEd_to_date.Mask = ""
    MaskEd_to_date.Text = ""
    MaskEd_to_date.Mask = "##/##/####"
    MsgBox "Please ENTER the  correct DATE."
    'MaskEd_to_date.SetFocus
    Exit Sub
End If
If rs_allreqflag = False Then
     RQ_ALLREQ1(0) = GET_REVDATE(MaskEd_from_date.FormattedText)
     RQ_ALLREQ1(1) = GET_REVDATE(MaskEd_to_date.FormattedText)
     RQ_ALLREQ1(2) = moda
     Set rs_allreq = RQ_ALLREQ1.OpenResultset(2, 3)
     rs_allreqflag = True
    
 Else
     RQ_ALLREQ1(0) = GET_REVDATE(MaskEd_from_date.FormattedText)
     RQ_ALLREQ1(1) = GET_REVDATE(MaskEd_to_date.FormattedText)
     RQ_ALLREQ1(2) = moda
     rs_allreq.MoreResults
     rs_allreq.Requery
End If
If rs_allreqflag2 = False Then
     RQ_ALLREQ2(0) = GET_REVDATE(MaskEd_from_date.FormattedText)
     RQ_ALLREQ2(1) = GET_REVDATE(MaskEd_to_date.FormattedText)
     RQ_ALLREQ2(2) = moda
     Set rs_allreq2 = RQ_ALLREQ2.OpenResultset(2, 3)
     rs_allreqflag2 = True
    
 Else
     RQ_ALLREQ2(0) = GET_REVDATE(MaskEd_from_date.FormattedText)
     RQ_ALLREQ2(1) = GET_REVDATE(MaskEd_to_date.FormattedText)
     RQ_ALLREQ2(2) = moda
     rs_allreq2.MoreResults
     rs_allreq2.Requery
End If


If opt_pr.Value = True Then
    If rs_allreq2.RowCount <> 0 Then
         cmd_print.Enabled = True
            rs_allreq2.MoveFirst
            
            msf_reqstat.Rows = 1
            srno = 1
      While Not rs_allreq2.EOF

            i = rs_allreq2.RowCount
                 
            'req_chglen (rs_allreq(0)) & vbTab & rs_allreq("CASENO") & vbTab & appt & vbTab & Memo & vbTab & payment & vbTab & EXAM & vbTab & REPORT
             msf_reqstat.Rows = msf_reqstat.Rows + 1
             msf_reqstat.row = srno
             msf_reqstat.col = 0
             msf_reqstat.TextMatrix(srno, 0) = srno
             msf_reqstat.TextMatrix(srno, 1) = Trim(req_chglen(rs_allreq2(0))) & "  " & Trim(rs_allreq2("reqndt"))
             
             If seldrequisitnrsflag = False Then
                seldrequisitn_qr(0) = rs_allreq2("reqnno")
                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                seldrequisitnrsflag = True
             Else
                 seldrequisitn_qr(0) = rs_allreq2("reqnno")
                seldrequisitn_rs.MoreResults
                seldrequisitn_rs.Requery
             End If
             If seldrequisitn_rs.RowCount > 0 Then
                msf_reqstat.TextMatrix(srno, 2) = seldrequisitn_rs("pathno")
             End If
             
             If selta101rsflag = False Then
                SELTA101_QR(0) = rs_allreq2("caseno")
                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                selta101rsflag = True
             Else
                SELTA101_QR(0) = rs_allreq2("caseno")
                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                SELTA101_RS.MoreResults
                SELTA101_RS.Requery
             End If
             If SELTA101_RS.RowCount > 0 Then
                msf_reqstat.TextMatrix(srno, 3) = rs_allreq2("CASENO")
                msf_reqstat.TextMatrix(srno, 4) = Trim(SELTA101_RS("tag")) & " " & Trim(SELTA101_RS("First_name")) & " " & Trim(SELTA101_RS("Mid_name")) & " " & Trim(SELTA101_RS("last_name"))
                msf_reqstat.TextMatrix(srno, 5) = Trim(SELTA101_RS("SEX"))
                msf_reqstat.TextMatrix(srno, 6) = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), rs_allreq2("REQNDT"))
             End If
             
            If seldrequisitn_rs.RowCount > 0 Then
                If seldrequisitn_rs.EOF = False Then
                msf_reqstat.TextMatrix(srno, 7) = " " & seldrequisitn_rs("NMETERIALS")
                msf_reqstat.ColWidth(7) = 0
                End If
                msf_reqstat.TextMatrix(srno, 9) = "" & seldrequisitn_rs("SITE")
            End If
            
            If SELDIMPRESIONFLAG = False Then
                seldimpresion_qr(0) = rs_allreq2("reqnno")
                Set seldimpresion_rs = seldimpresion_qr.OpenResultset(1, 3)
                SELDIMPRESIONFLAG = True
            Else
                seldimpresion_qr(0) = rs_allreq2("reqnno")
                seldimpresion_rs.MoreResults
                seldimpresion_rs.Requery
            End If
            If seldimpresion_rs.RowCount > 0 Then
                msf_reqstat.TextMatrix(srno, 8) = "" & seldimpresion_rs("impression")
            End If
            If rs_allreq2.RowCount > 0 Then
                If Trim(rs_allreq2("reqNstat")) = "T" Or Trim(rs_allreq2("reqNstat")) = "F" Then
                    msf_reqstat.TextMatrix(srno, 10) = "YES"
                     msf_reqstat.TextMatrix(srno, 11) = Trim(rs_allreq2("PREPORTBY"))
                     YES = YES + 1
                Else
                     msf_reqstat.TextMatrix(srno, 10) = "NO"
                     msf_reqstat.TextMatrix(srno, 11) = ""
                     NO = NO + 1
                End If
            End If
                rs_allreq2.MoveNext
               srno = srno + 1
            Wend
        '    MsgBox msf_reqstat.Rows
            LBL_RPTNO.Caption = "No. of Reports Done." & YES
            lbl_br.Caption = "No. of Reports not Done." & NO
        Else
        
        MsgBox "NO DATA is found between these two dates."
         'button_ena_dis ("cancel")
         'clear
         MaskEd_from_date.Enabled = False
         MaskEd_to_date.Enabled = False
         cmd_print.SetFocus
    End If
End If
If opt_pd.Value = True Then
    If rs_allreq.RowCount <> 0 Then
         cmd_print.Enabled = True
            rs_allreq.MoveFirst
            
            msf_reqstat.Rows = 1
            srno = 1
      While Not rs_allreq.EOF

            i = rs_allreq.RowCount
                 
            'req_chglen (rs_allreq(0)) & vbTab & rs_allreq("CASENO") & vbTab & appt & vbTab & Memo & vbTab & payment & vbTab & EXAM & vbTab & REPORT
             msf_reqstat.Rows = msf_reqstat.Rows + 1
             msf_reqstat.row = srno
             msf_reqstat.col = 0
             msf_reqstat.TextMatrix(srno, 0) = srno
             msf_reqstat.TextMatrix(srno, 1) = Trim(req_chglen(rs_allreq(0))) & "  " & Trim(rs_allreq("reqndt"))
             
             If seldrequisitnrsflag = False Then
                seldrequisitn_qr(0) = rs_allreq("reqnno")
                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                selta101rsflag = True
             Else
                 seldrequisitn_qr(0) = rs_allreq("reqnno")
                seldrequisitn_rs.MoreResults
                seldrequisitn_rs.Requery
             End If
             If seldrequisitn_rs.RowCount > 0 Then
                msf_reqstat.TextMatrix(srno, 2) = seldrequisitn_rs("pathno")
             End If
             
             If selta101rsflag = False Then
                SELTA101_QR(0) = rs_allreq("caseno")
                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                selta101rsflag = True
             Else
                SELTA101_QR(0) = rs_allreq("caseno")
                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                SELTA101_RS.MoreResults
                SELTA101_RS.Requery
             End If
             If SELTA101_RS.EOF Then
                If SELEMPBASINFO_FLAG = False Then
                    SELEMPBASINFO_QR(0) = Trim(rs_allreq("CASENO"))
                    Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
                    SELEMPBASINFO_FLAG = True
                Else
                    SELEMPBASINFO_QR(0) = Trim(rs_allreq("CASENO"))
                    SELEMPBASINFO_RS.MoreResults
                    SELEMPBASINFO_RS.Requery
                End If
                If SELEMPBASINFO_RS.RowCount > 0 Then
                   msf_reqstat.TextMatrix(srno, 3) = rs_allreq("CASENO")
                   msf_reqstat.TextMatrix(srno, 4) = Trim(SELEMPBASINFO_RS("TITLE")) & " " & Trim(SELEMPBASINFO_RS("Fname")) & " " & Trim(SELEMPBASINFO_RS("Mname")) & " " & Trim(SELEMPBASINFO_RS("lname"))
                   msf_reqstat.TextMatrix(srno, 5) = Trim(SELEMPBASINFO_RS("GENDER"))
                   msf_reqstat.TextMatrix(srno, 6) = DateDiff("YYYY", SELEMPBASINFO_RS("DOB"), rs_allreq("REQNDT"))
                End If
             End If
             If SELTA101_RS.RowCount > 0 Then
                msf_reqstat.TextMatrix(srno, 3) = rs_allreq("CASENO")
                msf_reqstat.TextMatrix(srno, 4) = Trim(SELTA101_RS("tag")) & " " & Trim(SELTA101_RS("First_name")) & " " & Trim(SELTA101_RS("Mid_name")) & " " & Trim(SELTA101_RS("last_name"))
                msf_reqstat.TextMatrix(srno, 5) = Trim(SELTA101_RS("SEX"))
                msf_reqstat.TextMatrix(srno, 6) = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), rs_allreq("REQNDT"))
             End If
             
            If seldrequisitn_rs.RowCount > 0 Then
                If seldrequisitn_rs.EOF = False Then
                msf_reqstat.TextMatrix(srno, 7) = " " & seldrequisitn_rs("NMETERIALS")
                msf_reqstat.ColWidth(7) = 0
                End If
                msf_reqstat.TextMatrix(srno, 9) = "" & seldrequisitn_rs("SITE")
            End If
            
            If SELDIMPRESIONFLAG = False Then
                seldimpresion_qr(0) = rs_allreq("reqnno")
                Set seldimpresion_rs = seldimpresion_qr.OpenResultset(1, 3)
                SELDIMPRESIONFLAG = True
            Else
                seldimpresion_qr(0) = rs_allreq("reqnno")
                seldimpresion_rs.MoreResults
                seldimpresion_rs.Requery
            End If
            If seldimpresion_rs.RowCount > 0 Then
                msf_reqstat.TextMatrix(srno, 8) = "" & seldimpresion_rs("impression")
            End If
            If rs_allreq.RowCount > 0 Then
                If Trim(rs_allreq("reqNstat")) = "T" Or Trim(rs_allreq("reqNstat")) = "F" Then
                    msf_reqstat.TextMatrix(srno, 10) = "YES"
                     msf_reqstat.TextMatrix(srno, 11) = Trim(rs_allreq("PREPORTBY"))
                     YES = YES + 1
                Else
                     msf_reqstat.TextMatrix(srno, 10) = "NO"
                     msf_reqstat.TextMatrix(srno, 11) = ""
                     NO = NO + 1
                End If
            End If
                rs_allreq.MoveNext
               srno = srno + 1
            Wend
        '    MsgBox msf_reqstat.Rows
            LBL_RPTNO.Caption = "No. of Reports Done." & YES
            lbl_br.Caption = "No. of Reports not Done." & NO
        Else
        
        MsgBox "NO DATA is found between these two dates."
         'button_ena_dis ("cancel")
         'clear
         MaskEd_from_date.Enabled = False
         MaskEd_to_date.Enabled = False
         cmd_print.SetFocus
    End If
End If

If OPT_NIP.Value = True Then
    vpp = 0
    lr = 0
    bar = 0
    drop1 = 0
    If SELRREQUISITNRSFLAG = False Then
        selrrequisitnqr(0) = moda
        selrrequisitnqr(1) = MaskEd_from_date.FormattedText
        selrrequisitnqr(2) = MaskEd_to_date.FormattedText
        Set selrrequisitnrs = selrrequisitnqr.OpenResultset(1, 3)
        SELRREQUISITNRSFLAG = True
    Else
        selrrequisitnqr(0) = moda
        selrrequisitnqr(1) = MaskEd_from_date.FormattedText
        selrrequisitnqr(2) = MaskEd_to_date.FormattedText
        selrrequisitnrs.MoreResults
        selrrequisitnrs.Requery
    End If
    If selrrequisitnrs.RowCount > 0 Then
         msf_reqstat.Rows = 1
         srno = 1
        
        While Not selrrequisitnrs.EOF
           
            
             msf_reqstat.Rows = msf_reqstat.Rows + 1
             msf_reqstat.row = srno
             msf_reqstat.col = 0
             msf_reqstat.TextMatrix(srno, 0) = srno
             msf_reqstat.TextMatrix(srno, 1) = Trim(req_chglen(selrrequisitnrs(0))) & "  " & Trim(selrrequisitnrs("reqndt"))
            
             If seldrequisitnrsflag = False Then
                seldrequisitn_qr(0) = selrrequisitnrs("reqnno")
                Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                selta101rsflag = True
             Else
                 seldrequisitn_qr(0) = selrrequisitnrs("reqnno")
                seldrequisitn_rs.MoreResults
                seldrequisitn_rs.Requery
             End If
             If seldrequisitn_rs.RowCount > 0 Then
                msf_reqstat.TextMatrix(srno, 2) = seldrequisitn_rs("pathno")
             End If
             
             If selta101rsflag = False Then
                SELTA101_QR(0) = selrrequisitnrs("caseno")
                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                selta101rsflag = True
             Else
                SELTA101_QR(0) = selrrequisitnrs("caseno")
                Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                SELTA101_RS.MoreResults
                SELTA101_RS.Requery
             End If
            
            If SELTA101_RS.RowCount > 0 Then
                msf_reqstat.TextMatrix(srno, 3) = selrrequisitnrs("CASENO")
                msf_reqstat.TextMatrix(srno, 4) = Trim(SELTA101_RS("tag")) & " " & Trim(SELTA101_RS("First_name")) & " " & Trim(SELTA101_RS("Mid_name")) & " " & Trim(SELTA101_RS("last_name"))
                msf_reqstat.TextMatrix(srno, 5) = Trim(SELTA101_RS("SEX"))
                msf_reqstat.TextMatrix(srno, 6) = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), selrrequisitnrs("REQNDT"))
             End If
            
             If seldrequisitn_rs.RowCount > 0 Then
                'If seldrequisitn_rs.EOF = False Then
                msf_reqstat.TextMatrix(srno, 7) = seldrequisitn_rs("REFDOC")
                'msf_reqstat.TextMatrix(srno, 8) = seldrequisitn_rs("REFTODOC")
                'End If
             End If
             
             If SELDIMPRESIONFLAG = False Then
                seldimpresion_qr(0) = selrrequisitnrs("reqnno")
                Set seldimpresion_rs = seldimpresion_qr.OpenResultset(1, 3)
                SELDIMPRESIONFLAG = True
            Else
                seldimpresion_qr(0) = selrrequisitnrs("reqnno")
                seldimpresion_rs.MoreResults
                seldimpresion_rs.Requery
                
            End If
            If seldimpresion_rs.RowCount > 0 Then
                msf_reqstat.TextMatrix(srno, 8) = seldimpresion_rs("IMPRESSION")
            End If
             If selrrequisitnrs.RowCount > 0 Then
                 If selrrequisitnrs("recordno") = 1 Then
                    msf_reqstat.col = 9
                    msf_reqstat.CellBackColor = vbGreen
                    msf_reqstat.TextMatrix(srno, 9) = "VPP"
                    
                    vpp = vpp + 1
                 End If
                 If selrrequisitnrs("recordno") = 2 Then
                    msf_reqstat.col = 9
                    msf_reqstat.CellBackColor = vbYellow
                    msf_reqstat.TextMatrix(srno, 9) = "LR"
                     
                    lr = lr + 1
                 End If
                 If selrrequisitnrs("recordno") = 3 Then
                    msf_reqstat.col = 9
                    msf_reqstat.CellBackColor = vbRed
                    msf_reqstat.TextMatrix(srno, 9) = "BAR"
                   
                    bar = bar + 1
                 End If
                 If selrrequisitnrs("recordno") = 4 Then
                    msf_reqstat.col = 9
                    msf_reqstat.CellBackColor = RGB(0, 255, 255)
                    msf_reqstat.TextMatrix(srno, 9) = "DrOP"
                    
                    drop1 = drop1 + 1
                 End If
                  If selrrequisitnrs.EOF = False Then
                    msf_reqstat.TextMatrix(srno, 10) = " " & selrrequisitnrs("FREPORTBY").Value
                  Else
                    msf_reqstat.TextMatrix(srno, 10) = " "
                  End If
                  msf_reqstat.TextMatrix(srno, 11) = selrrequisitnrs("RECPTNO")
             End If
           
                selrrequisitnrs.MoveNext
               srno = srno + 1
        Wend
            LBL_RPTNO.Caption = "No. of VPP Cases " & vpp
            lbl_lr.Caption = "No. or LR cases " & lr
            lbl_br.Caption = "No. or BAR cases " & bar
            lbl_drop.Caption = "No. or DrOp cases " & drop1
    End If
    
End If
If OPT_RFCASE.Value = True Then
    
      If SELRREQUISITN_RSFLAG = False Then
        SELRREQUISITN_QR(0) = MaskEd_from_date.FormattedText
        SELRREQUISITN_QR(1) = MaskEd_to_date.FormattedText
        SELRREQUISITN_QR(2) = moda
        Set SELRREQUISITN_RS = SELRREQUISITN_QR.OpenResultset(1, 3)
        SELRREQUISITN_RSFLAG = True
      Else
        SELRREQUISITN_QR(0) = MaskEd_from_date.FormattedText
        SELRREQUISITN_QR(1) = MaskEd_to_date.FormattedText
        SELRREQUISITN_QR(2) = moda
        SELRREQUISITN_RS.MoreResults
        SELRREQUISITN_RS.Requery
      End If
      If SELRREQUISITN_RS.RowCount > 0 Then
            LBL_RPTNO.Caption = "No. of RF. Cases is " & SELRREQUISITN_RS.RowCount
            msf_reqstat.Rows = 1
            srno = 1
            While Not SELRREQUISITN_RS.EOF
             msf_reqstat.Rows = msf_reqstat.Rows + 1
             msf_reqstat.row = srno
             msf_reqstat.col = 0
             msf_reqstat.TextMatrix(srno, 0) = srno
             msf_reqstat.TextMatrix(srno, 1) = Trim(req_chglen(SELRREQUISITN_RS(0))) & "  " & Trim(SELRREQUISITN_RS("reqndt"))
      
                If seldrequisitnrsflag = False Then
                   seldrequisitn_qr(0) = SELRREQUISITN_RS("reqnno")
                   Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                   selta101rsflag = True
                Else
                    seldrequisitn_qr(0) = SELRREQUISITN_RS("reqnno")
                   seldrequisitn_rs.MoreResults
                   seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                   msf_reqstat.TextMatrix(srno, 2) = seldrequisitn_rs("pathno")
                End If
             
                If selta101rsflag = False Then
                   SELTA101_QR(0) = SELRREQUISITN_RS("caseno")
                   Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                   selta101rsflag = True
                Else
                   SELTA101_QR(0) = SELRREQUISITN_RS("caseno")
                   Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                   SELTA101_RS.MoreResults
                   SELTA101_RS.Requery
                End If
            
            If SELTA101_RS.RowCount > 0 Then
                msf_reqstat.TextMatrix(srno, 3) = SELRREQUISITN_RS("CASENO")
                msf_reqstat.TextMatrix(srno, 4) = Trim(SELTA101_RS("tag")) & " " & Trim(SELTA101_RS("First_name")) & " " & Trim(SELTA101_RS("Mid_name")) & " " & Trim(SELTA101_RS("last_name"))
                msf_reqstat.TextMatrix(srno, 5) = Trim(SELTA101_RS("SEX"))
                msf_reqstat.TextMatrix(srno, 6) = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), SELRREQUISITN_RS("REQNDT"))
             End If
           SELRREQUISITN_RS.MoveNext
               srno = srno + 1
        Wend
      
      End If
End If
If OPT_DOCTORWISE.Value = True Then
    
    If SELFRREQUISITN_RSFLAG = False Then
        SELFRREQUISITN_QR(0) = MaskEd_from_date.FormattedText
        SELFRREQUISITN_QR(1) = MaskEd_to_date.FormattedText
        SELFRREQUISITN_QR(2) = moda
        SELFRREQUISITN_QR(3) = UCase(Trim(cmb_docname.Text))
        Set SELFRREQUISITN_RS = SELFRREQUISITN_QR.OpenResultset(1, 3)
        SELFRREQUISITN_RSFLAG = True
      Else
        SELFRREQUISITN_QR(0) = MaskEd_from_date.FormattedText
        SELFRREQUISITN_QR(1) = MaskEd_to_date.FormattedText
        SELFRREQUISITN_QR(2) = moda
        SELFRREQUISITN_QR(3) = UCase(Trim(cmb_docname.Text))
        SELFRREQUISITN_RS.MoreResults
        SELFRREQUISITN_RS.Requery
      End If
        If SELFRREQUISITN_RS.RowCount > 0 Then
            LBL_RPTNO.Caption = "No. of Patients " & SELFRREQUISITN_RS.RowCount
            msf_reqstat.Rows = 1
            srno = 1
            While Not SELFRREQUISITN_RS.EOF
             msf_reqstat.Rows = msf_reqstat.Rows + 1
             msf_reqstat.row = srno
             msf_reqstat.col = 0
             msf_reqstat.TextMatrix(srno, 0) = srno
             msf_reqstat.TextMatrix(srno, 1) = Trim(req_chglen(SELFRREQUISITN_RS(0))) & "  " & Trim(SELFRREQUISITN_RS("reqndt"))
      
                If seldrequisitnrsflag = False Then
                   seldrequisitn_qr(0) = SELFRREQUISITN_RS("reqnno")
                   Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
                   selta101rsflag = True
                Else
                    seldrequisitn_qr(0) = SELFRREQUISITN_RS("reqnno")
                   seldrequisitn_rs.MoreResults
                   seldrequisitn_rs.Requery
                End If
                If seldrequisitn_rs.RowCount > 0 Then
                   msf_reqstat.TextMatrix(srno, 2) = seldrequisitn_rs("pathno")
                End If
             
                If selta101rsflag = False Then
                   SELTA101_QR(0) = SELFRREQUISITN_RS("caseno")
                   Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                   selta101rsflag = True
                Else
                   SELTA101_QR(0) = SELFRREQUISITN_RS("caseno")
                   Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
                   SELTA101_RS.MoreResults
                   SELTA101_RS.Requery
                End If
            
            If SELTA101_RS.RowCount > 0 Then
                msf_reqstat.TextMatrix(srno, 3) = SELFRREQUISITN_RS("CASENO")
                msf_reqstat.TextMatrix(srno, 4) = Trim(SELTA101_RS("tag")) & " " & Trim(SELTA101_RS("First_name")) & " " & Trim(SELTA101_RS("Mid_name")) & " " & Trim(SELTA101_RS("last_name"))
                msf_reqstat.TextMatrix(srno, 5) = Trim(SELTA101_RS("SEX"))
                msf_reqstat.TextMatrix(srno, 6) = DateDiff("YYYY", SELTA101_RS("BIRTH_DATE"), SELFRREQUISITN_RS("REQNDT"))
             End If
             
             '****
             
             If SELFRREQUISITN_RS.RowCount > 0 Then
                 If SELFRREQUISITN_RS("recordno") = 0 Then
                    msf_reqstat.col = 7
                    'msf_reqstat.CellBackColor = vbGreen
                    msf_reqstat.TextMatrix(srno, 7) = ""
                    
                   
                 End If
                 If SELFRREQUISITN_RS("recordno") = 1 Then
                    msf_reqstat.col = 7
                    msf_reqstat.CellBackColor = vbGreen
                    msf_reqstat.TextMatrix(srno, 7) = "VPP"
                    
                    vpp = vpp + 1
                 End If
                 If SELFRREQUISITN_RS("recordno") = 2 Then
                    msf_reqstat.col = 7
                    msf_reqstat.CellBackColor = vbYellow
                    msf_reqstat.TextMatrix(srno, 7) = "LR"
                     
                    lr = lr + 1
                 End If
                 If SELFRREQUISITN_RS("recordno") = 3 Then
                    msf_reqstat.col = 7
                    msf_reqstat.CellBackColor = vbRed
                    msf_reqstat.TextMatrix(srno, 7) = "BAR"
                   
                    bar = bar + 1
                 End If
                 If SELFRREQUISITN_RS("recordno") = 4 Then
                    msf_reqstat.col = 7
                    msf_reqstat.CellBackColor = RGB(0, 255, 255)
                    msf_reqstat.TextMatrix(srno, 7) = "DrOP"
                    
                    drop1 = drop1 + 1
                 End If
'                  If SELFRREQUISITN_RS.EOF = False Then
'                    msf_reqstat.TextMatrix(srno, 8) = " " & SELFRREQUISITN_RS("FREPORTBY").Value
'                  Else
'                    msf_reqstat.TextMatrix(srno, 8) = " "
'                  End If
                  msf_reqstat.TextMatrix(srno, 8) = SELFRREQUISITN_RS("RECPTNO")
             End If
             
             '****
                SELFRREQUISITN_RS.MoveNext
               srno = srno + 1
        Wend
            LBL_RPTNO.Caption = "No. of VPP Cases " & vpp
            lbl_lr.Caption = "No. or LR cases " & lr
            lbl_br.Caption = "No. or BAR cases " & bar
            lbl_drop.Caption = "No. or DrOp cases " & drop1
      End If


End If
cmd_cancel.Enabled = True
cmd_cancel.SetFocus
Exit Sub
MSG:
MsgBox Err.Description

End Sub

Private Sub OPT_DOCTORWISE_Click()
If seldissecfinrsflag = False Then
    Set seldissecfin_rs = seldissecfin_qr.OpenResultset(1, 3)
    seldissecfinrsflag = True
Else
    seldissecfin_rs.MoreResults
    seldissecfin_rs.Requery
End If
If seldissecfin_rs.RowCount > 0 Then
    cmb_docname.clear
    While Not seldissecfin_rs.EOF
        cmb_docname.AddItem seldissecfin_rs("CCNAME")
        seldissecfin_rs.MoveNext
    Wend
End If
End Sub

Private Sub OPT_NIP_Click()
MaskEd_from_date.Text = DATE11
MaskEd_to_date.Text = DATE11

cmb_dept.SetFocus
End Sub

Private Sub opt_pd_Click()
MaskEd_from_date.Text = DATE11
MaskEd_to_date.Text = DATE11
cmb_dept.SetFocus
End Sub

Private Sub opt_pr_Click()
MaskEd_from_date.Text = DATE11
MaskEd_to_date.Text = DATE11
MaskEd_from_date.Enabled = True
MaskEd_to_date.Enabled = True
'cmb_dept.Text = Trim(SELMA106_RS1("DEPT_NAME"))
'Cmb_modality.Text = Trim(SELRMODALITY_RS1("MODLITYDES"))
cmb_dept.SetFocus
'MaskEd_from_date.SetFocus
End Sub

Private Sub OPT_RFCASE_Click()
MaskEd_from_date.Text = DATE11
MaskEd_to_date.Text = DATE11
MaskEd_from_date.Enabled = True
cmb_dept.SetFocus
End Sub
