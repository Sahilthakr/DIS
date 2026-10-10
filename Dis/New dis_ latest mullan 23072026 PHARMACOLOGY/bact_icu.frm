VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form bact_icu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14190
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
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
      Left            =   7320
      TabIndex        =   8
      Top             =   3240
      Width           =   1615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&CLEAR"
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
      Left            =   5520
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "E&XIT"
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
      Left            =   9120
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmd_print 
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
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.ComboBox cmb_ward 
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
      Left            =   5880
      TabIndex        =   1
      Top             =   1320
      Width           =   3735
   End
   Begin MSMask.MaskEdBox Msk_todt 
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
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
   Begin MSMask.MaskEdBox Msk_fromdt 
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   2040
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\disreportsec\microbiologyindex.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3120
      TabIndex        =   9
      Top             =   3000
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   4560
      TabIndex        =   10
      Top             =   1080
      Width           =   5295
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "WARD"
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
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TO DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FROM DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE,MULLANPUR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   13695
   End
   Begin VB.Label lbl_microbiology 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "BACTERIOLOGY - ISOLATED ORGANISMS REPORT "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   3480
      TabIndex        =   0
      Top             =   600
      Width           =   7275
   End
End
Attribute VB_Name = "bact_icu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Integer
Dim l As Integer
Dim sel_micro_qr As rdo.rdoQuery
Dim sel_micro_rs As rdo.rdoResultset
Dim str As String

Private Sub cmd_exit_Click()
End
End Sub

Private Sub cmd_print_Click()
If cmb_ward.Text = "" Then
    MsgBox "Please select ward"
    Exit Sub
End If

If msk_fromdt.Text = "__/__/____" Then
    MsgBox "Please select From date"
    msk_fromdt.SetFocus
    Exit Sub
End If

If msk_todt.Text = "__/__/____" Then
    MsgBox "Please select To date"
    msk_todt.SetFocus
    Exit Sub
End If

If Not IsDate(msk_fromdt.Text) Then
    MsgBox "Invalid Date (From date)", vbCritical
    msk_fromdt.SetFocus
    Exit Sub
End If

If Not IsDate(msk_todt.Text) Then
    MsgBox "Invalid Date (To Date)", vbCritical
    msk_todt.SetFocus
    Exit Sub
End If

Dim ward As String
var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)
ward = Trim(cmb_ward.Text)
If ward = "ALL" Then
    CrystalReport1.ReportFileName = "C:\disreportsec\bacteriology_icu_report.rpt"
    CrystalReport1.Formulas(0) = "FROMDT='" & Trim(msk_fromdt.FormattedText) & "'"
    CrystalReport1.Formulas(1) = "TODT='" & Trim(msk_todt.FormattedText) & "'"
    CrystalReport1.Formulas(2) = "WARD='" & Trim(cmb_ward.Text) & "'"
    CrystalReport1.SelectionFormula = "{RREQUISITN.Freportdt}>=DATE(" & var_date_f & ")AND {RREQUISITN.Freportdt}<=DATE(" & var_date_t & ") AND ({RREQUISITN.REQNSTAT}='F' or {RREQUISITN.REQNSTAT}='T') AND {RREQUISITN.DEPT_CODE}='GZZ'"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.Action = 1
Else
    CrystalReport1.ReportFileName = "C:\disreportsec\bacteriology_icu_report.rpt"
    CrystalReport1.Formulas(0) = "FROMDT='" & Trim(msk_fromdt.FormattedText) & "'"
    CrystalReport1.Formulas(1) = "TODT='" & Trim(msk_todt.FormattedText) & "'"
    CrystalReport1.Formulas(2) = "WARD='" & Trim(cmb_ward.Text) & "'"
    CrystalReport1.SelectionFormula = "{RREQUISITN.Freportdt}>=DATE(" & var_date_f & ")AND {RREQUISITN.Freportdt}<=DATE(" & var_date_t & ") AND ({RREQUISITN.REQNSTAT}='F' or {RREQUISITN.REQNSTAT}='T') AND {RREQUISITN.DEPT_CODE}='GZZ' and {DREQUISITN.WARD_NAME} = '" & ward & "'"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.Action = 1
End If
End Sub

Private Sub Command1_Click()
msk_fromdt.Text = "__/__/____"
msk_todt.Text = "__/__/____"
End Sub

Private Sub Command2_Click()
Me.Hide
DIS_MAIN.show
End Sub

Private Sub Form_Load()
DATE11 = Date
str = "SELECT distinct ward_name from  DISLIB.Drequisitn"
Set sel_micro_qr = rdb.CreateQuery("", str)
Set sel_micro_rs = sel_micro_qr.OpenResultset(1, 3)
sel_micro_rs.MoveFirst
i = 0
l = sel_micro_rs.RowCount
cmb_ward.AddItem ("ALL")
While (i < l)
    If IsNull(sel_micro_rs("ward_name")) = False Then
    cmb_ward.AddItem (sel_micro_rs("WARD_NAME"))
    End If
    sel_micro_rs.MoveNext
    i = i + 1
Wend
End Sub

Private Sub Msk_fromdt_GotFocus()
msk_fromdt.Text = DateAdd("d", -1, Format(DATE11, "dd/mm/yyyy"))
End Sub

Private Sub Msk_todt_GotFocus()
msk_todt.Text = Format(DATE11, "dd/mm/yyyy")
End Sub
