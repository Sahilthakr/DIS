VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12480
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
   ScaleHeight     =   8220
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR 
      Left            =   0
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   10
      Top             =   6360
      Width           =   6855
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
         Left            =   3600
         TabIndex        =   4
         Top             =   480
         Width           =   1215
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
         Left            =   480
         MouseIcon       =   "spquery.frx":0000
         TabIndex        =   0
         Top             =   480
         Width           =   1215
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
         Left            =   2040
         MouseIcon       =   "spquery.frx":030A
         TabIndex        =   3
         Top             =   480
         Width           =   1215
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
         Left            =   5160
         MouseIcon       =   "spquery.frx":0614
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FLEX 
      Height          =   3255
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "CLICK ON THE REQUISTION NO."
      Top             =   2760
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   11
      BackColor       =   16777215
      ForeColor       =   0
      ForeColorFixed  =   192
      WordWrap        =   -1  'True
      Enabled         =   -1  'True
      AllowUserResizing=   1
      FormatString    =   $"spquery.frx":091E
   End
   Begin MSMask.MaskEdBox msk_todt 
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   0
      Enabled         =   0   'False
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
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   4210752
      Enabled         =   0   'False
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "DEPARTMENT OF PATHOLOGY"
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
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   840
      Width           =   4455
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
      Left            =   2040
      TabIndex        =   8
      Top             =   1560
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
      Left            =   4680
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SELTA101_QR As rdoQuery
Dim SELTA101_RS As rdoResultset
Dim SELTA101FLAG As Boolean
Dim SELDREQ_QR As rdoQuery
Dim SELDREQ_RS As rdoResultset
Dim SELDREQFLAG As Boolean
Dim SELRREQ_QR As rdoQuery
Dim SELRREQ_RS As rdoResultset
Dim SELRREQFLAG As Boolean
Dim SELMA301_QR As rdoQuery
Dim SELMA301_RS As rdoResultset
Dim SELMA301_FLAG As Boolean
Dim SELDREQ_QR1 As rdoQuery
Dim SELDREQ_RS1 As rdoResultset
Dim SELDREQ_FLAG As Boolean
Dim RS As rdoResultset
Dim II As Boolean
Dim JJ As Boolean
Public DFD
Public DTD
Public K
Private Sub cmd_back_Click()
    msk_fromdt.Mask = ""
    msk_fromdt.Text = ""
    msk_fromdt.Mask = "##/##/####"
    msk_todt.Mask = ""
    msk_todt.Text = ""
    msk_todt.Mask = "##/##/####"
    'Me.MousePointer = 11
    DIS_MAIN.show
    Me.Hide
    'Me.MousePointer = 0
End Sub

Private Sub cmd_cancel_Click()
msk_fromdt.Text = "__/__/____"
msk_todt.Text = "__/__/____"
FLEX.Rows = 1
msk_fromdt.Enabled = False
msk_todt.Enabled = False
cmd_view.Enabled = True
End Sub

Private Sub cmd_print_Click()
DFD = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
DTD = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
CR.ReportFileName = "D:\dissource\DISREPORT\PATHOLOGY.rpt"
CR.SelectionFormula = "{RREQUISITN.REQNDT}>=DATE(" & DFD & ")AND {RREQUISITN.REQNDT}<=DATE(" & DTD & ")"
CR.Formulas(0) = "FRMDT= '" & msk_fromdt.Text & "'"
CR.Formulas(1) = "TODT='" & msk_todt.Text & "'"
CR.Action = 1

End Sub

Private Sub cmd_view_Click()
msk_fromdt.Enabled = True
msk_todt.Enabled = True
cmd_view.Enabled = False
msk_fromdt.SetFocus
End Sub

Private Sub Form_Load()
II = True
JJ = True
str = "SELECT CASE_NO,TAG,FIRST_NAME,MID_NAME,LAST_NAME,SEX,UD_NAME FROM PATOOLSLIB.TA101 WHERE CASE_NO=?"
Set SELTA101_QR = rdb.CreateQuery("TA", str)
Set RS = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107") '***2/2/99
DATE11 = RS("dt1") '**add 24/11/2k
str = "SELECT PATHNO,NMETERIALS FROM DISLIB.DREQUISITN WHERE CASENO=? AND REQNO LIKE '%SP%'"
Set SELDREQ_QR = rdb.CreateQuery("DREQ", str)
str = "SELECT REQNDT,CASENO,DEPT_CODE,MODALITYCD,REFBYDOC,FREPORTBY,FREPORTDT,RECPTDT,REQNSTAT FROM RISLIB.RREQUISITN WHERE REQNDT BETWEEN ? AND ? AND CASENO LIKE '%RF%'AND MODALITYCD='SP' AND DEPT_CODE='FZZ'AND (REQNSTAT='T'OR REQNSTAT='F') ORDER BY REQNDT DESC"
Set SELRREQ_QR = rdb.CreateQuery("RREQ", str)
str = "SELECT SITE FROM DISLIB.DREQUISITN WHERE PATHNO=?"
Set SELDREQ_QR1 = rdb.CreateQuery("DREQ1", str)
str = "SELECT SITEDESC FROM PATOOLSLIB.MA301 WHERE SITECD=? "
Set SELMA301_QR = rdb.CreateQuery("MA301U", str)
End Sub

Private Sub Msk_fromdt_GotFocus()
    msk_fromdt.Text = Format(DATE11, "DD/MM/YYYY")
End Sub

Private Sub Msk_todt_GotFocus()
    msk_todt.Text = Format(DATE11, "DD/MM/YYYY")
End Sub

Private Sub msk_todt_LostFocus()
On Error GoTo MSG:
If InStr(msk_todt.Text, "_") > 0 Then
   msk_todt.Mask = ""
   msk_todt.Text = ""
   msk_todt.Mask = "##/##/####"
   Exit Sub
End If
If IsDate(GET_REVDATE(msk_todt.Text)) = False Then
    msk_todt.Mask = ""
    msk_todt.Text = ""
    msk_todt.Mask = "##/##/####"
    MsgBox "Please ENTER the  correct DATE."
    msk_todt.SetFocus
    Exit Sub
End If
msk_todt.Enabled = False
msk_fromdt.Enabled = False
SELRREQ_QR(0) = GET_REVDATE(msk_fromdt.Text)
SELRREQ_QR(1) = GET_REVDATE(msk_todt.Text)

Dim str
Dim i
i = 1
FLEX.Rows = 2
Set SELRREQ_RS = SELRREQ_QR.OpenResultset(1, 3)
l:
'DoEvents
    While Not SELRREQ_RS.EOF
    JJ = False
   
    Me.MousePointer = vbHourglass
    If SELRREQ_RS("CASENO") = K Then
        SELRREQ_RS.MoveNext
        GoTo l:
    End If
    FLEX.TextMatrix(i, 0) = i
    FLEX.TextMatrix(i, 2) = SELRREQ_RS("CASENO")
    K = SELRREQ_RS("CASENO")
    If IsNull(SELRREQ_RS("RECPTDT")) Then
        FLEX.TextMatrix(i, 7) = ""
        Else
        FLEX.TextMatrix(i, 7) = SELRREQ_RS("RECPTDT")
    End If
    FLEX.TextMatrix(i, 10) = SELRREQ_RS("REFBYDOC")
    FLEX.ColAlignment(10) = flexAlignLeftCenter
    If IsNull(SELRREQ_RS("FREPORTDT")) Then
        FLEX.TextMatrix(i, 8) = ""
        Else
        FLEX.TextMatrix(i, 8) = SELRREQ_RS("FREPORTDT")
    End If
    If IsNull(SELRREQ_RS("FREPORTBY")) Then
        FLEX.TextMatrix(i, 9) = ""
        Else
        FLEX.TextMatrix(i, 9) = SELRREQ_RS("FREPORTBY")
        FLEX.ColAlignment(9) = flexAlignLeftCenter
    End If
    SELTA101_QR(0) = FLEX.TextMatrix(i, 2)
    Set SELTA101_RS = SELTA101_QR.OpenResultset(1, 3)
    JJ = True
    While Not SELTA101_RS.EOF
            JJ = False
            FLEX.TextMatrix(i, 3) = Trim(SELTA101_RS("TAG")) & " " & Trim(SELTA101_RS("FIRST_NAME")) & " " & Trim(SELTA101_RS("MID_NAME")) & " " & Trim(SELTA101_RS("LAST_NAME"))
            FLEX.TextMatrix(i, 4) = SELTA101_RS("SEX")
            FLEX.ColAlignment(3) = flexAlignLeftCenter
            SELTA101_RS.MoveNext
    Wend
    SELDREQ_QR(0) = FLEX.TextMatrix(i, 2)
    Set SELDREQ_RS = SELDREQ_QR.OpenResultset(1, 3)
    JJ = True
    While Not SELDREQ_RS.EOF
        JJ = False
        If II = False Then
            i = i + 1
            FLEX.Rows = i + 1
            FLEX.TextMatrix(i, 0) = i
            FLEX.TextMatrix(i, 2) = SELRREQ_RS("CASENO")
            FLEX.TextMatrix(i, 3) = FLEX.TextMatrix(i - 1, 3)
            FLEX.TextMatrix(i, 4) = FLEX.TextMatrix(i - 1, 4)
            If IsNull(SELRREQ_RS("RECPTDT")) Then
                FLEX.TextMatrix(i, 7) = ""
                Else
                FLEX.TextMatrix(i, 7) = SELRREQ_RS("RECPTDT")
            End If
            FLEX.TextMatrix(i, 10) = SELRREQ_RS("REFBYDOC")
            FLEX.ColAlignment(10) = flexAlignLeftCenter
            If IsNull(SELRREQ_RS("FREPORTDT")) Then
                FLEX.TextMatrix(i, 8) = ""
                Else
                FLEX.TextMatrix(i, 8) = SELRREQ_RS("FREPORTDT")
            End If
            If IsNull(SELRREQ_RS("FREPORTBY")) Then
                FLEX.TextMatrix(i, 9) = ""
                Else
                FLEX.TextMatrix(i, 9) = SELRREQ_RS("FREPORTBY")
                FLEX.ColAlignment(9) = flexAlignLeftCenter
            End If
        End If
        FLEX.TextMatrix(i, 1) = SELDREQ_RS("PATHNO")
        If IsNull(SELDREQ_RS("NMETERIALS")) Then
            FLEX.TextMatrix(i, 5) = ""
            Else
            FLEX.TextMatrix(i, 5) = SELDREQ_RS("NMETERIALS")
        End If
        FLEX.ColAlignment(5) = flexAlignLeftCenter
        SELDREQ_RS.MoveNext
        II = False
    Wend
    SELDREQ_QR1(0) = FLEX.TextMatrix(i, 1)
    If Trim(SELDREQ_QR1(0)) = "" Then
        FLEX.TextMatrix(i, 6) = ""
    Else
    Set SELDREQ_RS1 = SELDREQ_QR1.OpenResultset(1, 3)
    JJ = True
        If II = False Then
        'i = i + 1
        While Not SELDREQ_RS1.EOF
            JJ = False
            SITECD = SELDREQ_RS1("SITE")
            SELMA301_QR(0) = SITECD
            Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
            JJ = True
            While Not SELMA301_RS.EOF
                JJ = False
                FLEX.TextMatrix(i, 6) = SELMA301_RS("SITEDESC")
                SELMA301_RS.MoveNext
            Wend
            II = False
            SELDREQ_RS1.MoveNext
        Wend
        End If
    End If
    II = True
    SELRREQ_RS.MoveNext
    i = i + 1
    FLEX.Rows = i + 1
    
Wend
    Me.MousePointer = vbDefault
    If JJ = False Then
        SELTA101_RS.MoreResults
        SELTA101_RS.Requery
    End If
    If JJ = False Then
        SELDREQ_RS.MoreResults
        SELDREQ_RS.Requery
    End If
    If JJ = False Then
        SELRREQ_RS.MoreResults
        SELRREQ_RS.Requery
        SELDREQ_RS1.MoreResults
        SELDREQ_RS1.Requery
        SELMA301_RS.MoreResults
        SELMA301_RS.Requery
    End If
MSG:
MsgBox Err.Description
End Sub
