VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_CBCedit 
   BackColor       =   &H00C0C0C0&
   Caption         =   "CBC Edit window"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13890
   LinkTopic       =   "Form2"
   ScaleHeight     =   10590
   ScaleWidth      =   13890
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1320
      Top             =   840
   End
   Begin VB.TextBox TXT_OPERATOR 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   4320
      TabIndex        =   25
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton cmd_mdc 
      BackColor       =   &H80000003&
      Caption         =   "&MDC"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmd_back 
      BackColor       =   &H80000003&
      Caption         =   "&BACK"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9960
      Width           =   1335
   End
   Begin VB.CommandButton cmd_crossmatch 
      BackColor       =   &H80000003&
      Caption         =   "C&ROSSMATCH"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmd_commit 
      BackColor       =   &H80000003&
      Caption         =   "&COMMIT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9960
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   9735
      Begin VB.TextBox txt_hinter 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7560
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   1800
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6015
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   10610
         _Version        =   393216
         Rows            =   4
         FixedRows       =   3
         BackColor       =   16777215
         BackColorFixed  =   -2147483645
         BackColorBkg    =   16777215
         GridColor       =   8421504
         GridLines       =   3
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H80000006&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000009&
         Height          =   6255
         Left            =   120
         Top             =   120
         Width           =   9495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   9735
      Begin VB.TextBox txt_provdiag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         TabIndex        =   21
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txt_pridoc 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   19
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txt_category 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txt_age 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   5655
      End
      Begin MSMask.MaskEdBox msk_caseno 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin VB.Label lbl_ProvDiag 
         BackStyle       =   0  'Transparent
         Caption         =   "Provisional Diagnosis"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   20
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lbl_pridoc 
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Doctor/Service"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lbl_category 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lbl_age 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3720
         TabIndex        =   7
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Case No"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000009&
         Height          =   1455
         Left            =   120
         Top             =   120
         Width           =   9495
      End
   End
   Begin MSMask.MaskEdBox msk_sysdate 
      Height          =   315
      Left            =   10440
      TabIndex        =   27
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
   Begin MSMask.MaskEdBox msk_systime 
      Height          =   315
      Left            =   10440
      TabIndex        =   28
      Top             =   600
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm:ss"
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   480
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   0
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport3 
      Left            =   840
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   30
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   29
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "OPERATED BY"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   3000
      TabIndex        =   15
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label Label5 
      Caption         =   "2/11/21"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      Height          =   615
      Left            =   3120
      Top             =   9840
      Width           =   3135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1170
      TabIndex        =   11
      Top             =   120
      Width           =   7770
   End
End
Attribute VB_Name = "frm_CBCedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sel_rreq_qr As rdoQuery
Dim sel_rreq_rs As rdoResultset

Dim SEL_TA101_QR As rdoQuery
Dim SEL_TA101_RS As rdoResultset

Dim sel_dhemcbc_qr As rdoQuery
Dim sel_dhemcbc_rs As rdoResultset

Dim sel_editHINT_qr As rdoQuery
Dim sel_editHINT_rs As rdoResultset

Dim sel_ta1012_qr As rdoQuery
Dim sel_ta1012_rs As rdoResultset

Dim R_NO As String
'***** DHEMCBCDET insertion *****
Dim SELCASENO_QR As rdoQuery
Dim SELCASENO_RS As rdoResultset

Dim SELSEXAGE_QR As rdoQuery
Dim SELSEXAGE_RS As rdoResultset

Dim SELSTAFF_QR As rdoQuery
Dim SELSTAFF_RS As rdoResultset

Dim SELHINTFVAL_QR As rdoQuery
Dim SELHINTFVAL_RS As rdoResultset

Dim INSDHEMCBCDET_QR As rdoQuery

Dim INSCBCINTFERR_QR As rdoQuery

Dim SELDHEMCBCQR As rdoQuery
Dim SELDHEMCBCRS As rdoResultset

Dim SELDHEMCBCQRFLAG As Boolean

Dim HAGE As String
Dim HBSEX As String

Dim SA, age, SEX As String

Dim updhinterface_qr As rdoQuery

Dim updrrequisitn1_qr As rdoQuery

Dim UPDRTIMEC_QR As rdoQuery

Dim USRNAME_QR As rdoQuery
Dim USRNAME_RS As rdoResultset
Dim LOGINUSRNAME As String
Dim LOGINTYP As String
Dim LOGINUSRID As String

Dim CASENO As String

Dim SELPATNSTAT_QR As rdoQuery
Dim SELPATNSTAT_RS As rdoResultset

Dim sel_dhem_qr As rdoQuery
Dim sel_dhem_rs As rdoResultset


'********************************

Dim l_date As String
Dim f_date

Private Sub cmd_back_Click()
txt_hinter.Visible = False
Unload Me
'Me.Hide
frm_CBCworklist.show
End Sub

Private Sub cmd_cancel_Click()
    MSFlexGrid1.clear
    txt_hinter.Text = ""
End Sub

Private Sub cmd_crossmatch_Click()
'CrystalReport3.SelectionFormula = "{BLOODPROLF.REQNO} = '" & req_chglen(Trim(cmb_reqno.Text)) & "'"
'CrystalReport3.ReportFileName = "C:\DISREPORTSEC\DISCOMRP.RPT"
'CrystalReport3.Formulas(0) = "NAME = '" & Trim(txt_name.Text) & "'"
'CrystalReport3.Formulas(1) = "CASENO = '" & Trim(MaskEdBox1.FormattedText) & "'"
'CrystalReport3.Formulas(2) = "GRP = '" & Trim(CMB_BG.Text) & "'"
'CrystalReport3.Formulas(3) = "ent = '" & Trim(cmb_enterby.Text) & "'"
'CrystalReport3.Formulas(4) = "Dtmno = '" & Trim(txt_pathno.Text) & "'"
'CrystalReport3.Formulas(5) = "Udname = '" & Trim(txt_refdoc.Text) & "'"
'CrystalReport3.Formulas(6) = "Category = '" & Trim(txt_category.Text) & "'"
'wb1 = ""
'wb1 = Trim(txt_ward.Text) & " / " & Trim(txt_bedno.Text)
'CrystalReport3.Formulas(7) = "WB = '" & Trim(wb1) & "'"
'CrystalReport3.Formulas(8) = "Clindiag = '" & txt_relclnhis.Text & "'"
'CrystalReport3.WindowState = crptMaximized
'CrystalReport3.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
'CrystalReport3.Action = 1

Dim sel_rrequisitn_qr As rdoQuery
Dim sel_rrequisitn_rs As rdoResultset

Dim SEL_BLOODPROLF_QR As rdoQuery
Dim SEL_BLOODPROLF_RS As rdoResultset

Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset

str = "SELECT * from rislib.rrequisitn where caseno=? and DEPT_CODE='NZZ' AND REQNSTAT='T' and MODALITYCD='BP' order by REQNNO desc"
Set sel_rrequisitn_qr = rdb.CreateQuery("", str)

sel_rrequisitn_qr(0) = Trim(msk_caseno.FormattedText)
Set sel_rrequisitn_rs = sel_rrequisitn_qr.OpenResultset(1, 3)

If Not sel_rrequisitn_rs.EOF Then
    

    str = "select * from dislib.BLOODPROLF where reqno =? and (compcode='PC' or compcode='WB' or compcode='pc' or compcode='wb')"
    Set SEL_BLOODPROLF_QR = rdb.CreateQuery("", str)

    SEL_BLOODPROLF_QR(0) = Trim(sel_rrequisitn_rs("reqnno"))
    Set SEL_BLOODPROLF_RS = SEL_BLOODPROLF_QR.OpenResultset(1, 3)
    
    If Not SEL_BLOODPROLF_RS.EOF Then
        '***** Report related coding *****
        MsgBox "yes"
        
        CrystalReport3.SelectionFormula = "{BLOODPROLF.REQNO} = '" & Trim(sel_rrequisitn_rs("reqnno")) & "'"
        CrystalReport3.ReportFileName = "C:\DISREPORTSEC\DISCOMRP.RPT"
        
        Dim SELDTMNO_QR As rdoQuery
        Dim SELDTMNO_RS As rdoResultset
        
        str = "SELECT * FROM DISLIB.DTMNO WHERE CASENO=?"
        Set SELDTMNO_QR = rdb.CreateQuery("", str)
        
        SELDTMNO_QR(0) = Trim(msk_caseno.FormattedText)
        Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
        
        str = "select * from dislib.drequisitn where reqno=?"
        Set seldrequisitn_qr = rdb.CreateQuery("", str)
        
        seldrequisitn_qr(0) = Trim(sel_rrequisitn_rs("reqnno"))
        Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
        
        CrystalReport3.Formulas(0) = "NAME = '" & Trim(txt_name.Text) & "'"
        CrystalReport3.Formulas(1) = "CASENO = '" & Trim(msk_caseno.FormattedText) & "'"
        If Not SELDTMNO_RS.EOF Then
            If Trim(SELDTMNO_RS("FLD1")) <> "" Then
                CrystalReport3.Formulas(2) = "GRP = '" & Trim(SELDTMNO_RS("FLD1")) & "'"
            Else
                CrystalReport3.Formulas(2) = "GRP = ''"
            End If
            If Trim(SELDTMNO_RS("DTMNO")) <> "" Then
                CrystalReport3.Formulas(4) = "Dtmno = '" & Trim(SELDTMNO_RS("DTMNO")) & "'"
            Else
                CrystalReport3.Formulas(4) = "Dtmno = ''"
            End If
        Else
            CrystalReport3.Formulas(2) = "GRP = ''"
            CrystalReport3.Formulas(4) = "Dtmno = ''"
        End If
        If Trim(sel_rrequisitn_rs("ENTERBY")) <> "" Then
            CrystalReport3.Formulas(3) = "ent = '" & Trim(sel_rrequisitn_rs("ENTERBY")) & "'"
        Else
            CrystalReport3.Formulas(3) = "ent = ''"
        End If
        If Trim(sel_rrequisitn_rs("REFBYDOC")) <> "" Then
            CrystalReport3.Formulas(5) = "Udname = '" & Trim(sel_rrequisitn_rs("REFBYDOC")) & "'"
        Else
            CrystalReport3.Formulas(5) = "Udname = ''"
        End If
        
        CrystalReport3.Formulas(6) = "Category = '" & Trim(txt_catEgory.Text) & "'"
        If Not seldrequisitn_rs.EOF Then
            If Trim(seldrequisitn_rs("WARD_NAME")) <> "" And Trim(seldrequisitn_rs("bedno")) Then
                CrystalReport2.Formulas(7) = "WB = '" & Trim(seldrequisitn_rs("WARD_NAME")) & "/" & Trim(seldrequisitn_rs("bedno")) & "'"
            Else
                CrystalReport2.Formulas(7) = "WB = ''"
            End If
        Else
            CrystalReport2.Formulas(7) = "WB = ''"
        End If
        If Trim(seldrequisitn_rs("NMETERIALS")) <> "" Then
            CrystalReport3.Formulas(8) = "Clindiag = '" & Trim(seldrequisitn_rs("NMETERIALS")) & "'"
        Else
            CrystalReport3.Formulas(8) = "Clindiag = ''"
        End If
        
        CrystalReport3.WindowState = crptMaximized
        CrystalReport3.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport3.Action = 1
        '*********************************
    Else
        MsgBox "No Blood Product report found for this Case No"
    End If
Else
    MsgBox "No Blood Product report found for this Case No"
End If



End Sub

Private Sub cmd_mdc_Click()

str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1

Dim flex_Cols
Dim SEL_DCBCREP1_QR As rdoQuery
Dim SEL_DCBCREP1_RS As rdoResultset
flex_Cols = MSFlexGrid1.Cols
For i = 2 To flex_Cols - 1

    If MSFlexGrid1.TextMatrix(0, i) <> "" Then
        
        'Check for entry in select * from dislib/dcbcrep1 where reqno='UZZHL09233832'
        
        str = "select * from dislib.dcbcrep1 where reqno =?"
        Set SEL_DCBCREP1_QR = rdb.CreateQuery("", str)
        
        SEL_DCBCREP1_QR(0) = MSFlexGrid1.TextMatrix(0, i)
        Set SEL_DCBCREP1_RS = SEL_DCBCREP1_QR.OpenResultset(1, 1)
        
        If Not SEL_DCBCREP1_RS.EOF Then
            CrystalReport2.SelectionFormula = "{RREQUISITN.REQNNO} = '" & MSFlexGrid1.TextMatrix(0, i) & "'"
            GoTo MDCstart
        Else
            'Exit Sub
        End If
    Else
        MsgBox "No MDC report found for this Case No"
        Exit Sub
    End If

Next

If CrystalReport2.SelectionFormula = "" Then
    MsgBox "No MDC report found for this Case No"
    Exit Sub
End If

MDCstart:
CrystalReport2.ReportFileName = "C:\DISREPORTSEC\HEMATO.RPT"
CrystalReport2.Formulas(0) = "REQNO='" & Trim(R_NO) & "'"
CrystalReport2.WindowState = crptMaximized
CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport2.Action = 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

'str = "select DISTINCT(REQNNO),freportdt,REMARKS,PRO_CODE, OBSERVALUE  from rislib.rrequisitn A, DISLIB.DHEMCBCDET B where A.REQNNO=B.REQNO AND CASENO=? and MODALITYCD='HL' and FREPORTDT>=? order by  reqnno desc fetch first 3 rows only"
'PRO_CODE, OBSERVALUE added on 24/11/09
str = "select DISTINCT(REQNNO),freportdt,REMARKS from rislib.rrequisitn A, DISLIB.DHEMCBCDET B where A.REQNNO=B.REQNO AND CASENO=? and MODALITYCD='HL' and FREPORTDT>=? order by  reqnno desc fetch first 3 rows only"
'str = "select * from rislib.rrequisitn where caseno=? and MODALITYCD='HL' and FREPORTDT>=? order by caseno, reqnno desc fetch first 3 rows only"
Set sel_rreq_qr = rdb.CreateQuery("", str)

str = "select * from PATOOLSLIB.ta101 where case_no =?"
Set SEL_TA101_QR = rdb.CreateQuery("", str)

str = "select * from DISLIB.DHEMCBCDET WHERE REQNO=? AND PRO_CODE=?"
Set sel_dhemcbc_qr = rdb.CreateQuery("", str)

'***** Insert into DHEMCBCDET Procedure *****
str = "SELECT CASENO,PATNSTAT FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELCASENO_QR = rdb.CreateQuery("", str)

str = "SELECT SEX,BIRTH_DATE FROM PATOOLSLIB.TA101 WHERE CASE_NO=?"
Set SELSEXAGE_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=? "
Set SELSTAFF_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM INTERFACE.HINTERFACE WHERE REQNO=?"
Set SELHINTFVAL_QR = rdb.CreateQuery("", str)

str = "INSERT INTO DISLIB.DHEMCBCDET (REQNO,PRO_CODE,OBSERVALUE,OBSER_TYPE,SOURCE,ENTRYBY,REMARKS,UOM,NRANGE,PVALUE) VALUES(?,?,?,?,?,?,?,?,?,?)"
Set INSDHEMCBCDET_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DHEMCBCRNG WHERE PROCODE=? AND AGE2=? AND SEX=? "
Set SELDHEMCBCQR = rdb.CreateQuery("", str)

str = "INSERT INTO INTERFACE.CBCINTFERR (REQNO,ERRORCODE,ERRORMODE,ERRORDESC,ERRORDATE,ERRORTIME,PRINTFLAG) VALUES(?,?,?,?,?,?,?)"
Set INSCBCINTFERR_QR = rdb.CreateQuery("", str)

str = "select * from DISLIB.DHEMCBCDET where REQNO =?"
Set sel_dhem_qr = rdb.CreateQuery("", str)



'Updates
str = "update INTERFACE.hinterface set iflag = '3' where reqno = ?"
Set updhinterface_qr = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
Set updrrequisitn1_qr = rdb.CreateQuery("updrrequisitn1_qr1NHE", str)

str = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=?,PRNREPDATE=?,PRNREPTIME=? WHERE REQNNO=?"
Set UPDRTIMEC_QR = rdb.CreateQuery("QWEQWEQWEccy", str)


'SECURITY STARTS
msk_sysdate.Text = Format(DATE11, "DD/MM/YYYY")
msk_systime.Text = Format(Time, "HH:MM:SS")
    
str = "SELECT * FROM DISLIB.DISLNKUSR WHERE USR_ID = ? "
Set USRNAME_QR = rdb.CreateQuery("USRNAME_QR", str)

logonfrm.Main
datetime
USRNAME_QR(0) = LOGIN
'USRNAME_QR(0) = "A" + logonfrm.txtusername.Text
Set USRNAME_RS = USRNAME_QR.OpenResultset(1, 1)

If USRNAME_RS.RowCount > 0 Then
    LOGINUSRNAME = Trim(USRNAME_RS("USERNAME"))
    LOGINTYP = Trim(USRNAME_RS("USRTYPE"))
    LOGINUSRID = Trim(USRNAME_RS("usr_id"))
    TXT_OPERATOR.Text = (LOGINUSRNAME)
    TXT_OPERATOR.Enabled = False
Else
    MsgBox "The Security Details of this User is not available!!", vbInformation, "DIS"
    Exit Sub
End If
        
'SECURITY ENDS

CBCrec = False
'panic = False
'Print
str = "SELECT PATNSTAT FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELPATNSTAT_QR = rdb.CreateQuery("", str)
'********************************************
End Sub
Private Sub Timer1_Timer()
    datetime
    msk_sysdate.Text = Format(DATE11, "DD/MM/YYYY")
    msk_systime.Text = Format(time1, "HH:MM:SS")
End Sub

Public Function fillData(cno As String, rno As String)




Call Get_Hinterface(cno)
Cal_Abs_values
CASENO = cno
R_NO = rno

MSFlexGrid1.Rows = 23
MSFlexGrid1.Cols = 5
MSFlexGrid1.ColWidth(0) = 1800

MSFlexGrid1.TextMatrix(0, 0) = "REQUISITN NO."
MSFlexGrid1.TextMatrix(1, 0) = "REPORT DATE"
MSFlexGrid1.TextMatrix(2, 0) = "COUNTER NAME"
MSFlexGrid1.TextMatrix(3, 0) = "HAEMOGLOBIN"
MSFlexGrid1.TextMatrix(4, 0) = "RBC COUNT"
MSFlexGrid1.TextMatrix(5, 0) = "HAEMATOCRIT"
MSFlexGrid1.TextMatrix(6, 0) = "MCV"
MSFlexGrid1.TextMatrix(7, 0) = "MCH"
MSFlexGrid1.TextMatrix(8, 0) = "MCHC"
MSFlexGrid1.TextMatrix(9, 0) = "RDW"
MSFlexGrid1.TextMatrix(10, 0) = "PLATELETS"
MSFlexGrid1.TextMatrix(11, 0) = "MPV"
MSFlexGrid1.TextMatrix(12, 0) = "WBC COUNT"
MSFlexGrid1.TextMatrix(13, 0) = "NEUTROPHIL %"
MSFlexGrid1.TextMatrix(14, 0) = "LYMPHOCYTE %"
MSFlexGrid1.TextMatrix(15, 0) = "MONOCYTE %"
MSFlexGrid1.TextMatrix(16, 0) = "EOSINOPHIL %"
MSFlexGrid1.TextMatrix(17, 0) = "BASOPHIL %"
MSFlexGrid1.TextMatrix(18, 0) = "NEUTROPHIL ABS"
MSFlexGrid1.TextMatrix(19, 0) = "LYMPHOCYTE ABS"
MSFlexGrid1.TextMatrix(20, 0) = "MONOCYTE ABS"
MSFlexGrid1.TextMatrix(21, 0) = "EOSINOPHIL ABS"
MSFlexGrid1.TextMatrix(22, 0) = "BASOPHIL ABS"

'***** Upper Panel *****

SEL_TA101_QR(0) = cno
Set SEL_TA101_RS = SEL_TA101_QR.OpenResultset(1, 3)

msk_caseno.Text = cno
txt_sex.Text = Trim(SEL_TA101_RS("sex"))
txt_name.Text = Trim(SEL_TA101_RS("first_name")) & " " & Trim(SEL_TA101_RS("mid_name")) & " " & Trim(SEL_TA101_RS("last_name"))
txt_Age = DateDiff("YYYY", SEL_TA101_RS("BIRTH_DATE"), CStr(DATE11))
                           
If Trim(txt_Age.Text) < 12 Then
    Dim age1 As String
    age1 = DateDiff("m", SEL_TA101_RS("BIRTH_DATE"), CStr(DATE11))
    age1 = "." & (age Mod 12)
    If age1 = ".0" Then
        age1 = DateDiff("d", SEL_TA101_RS("BIRTH_DATE"), CStr(DATE11))
        age1 = ".0" & (age1 Mod 30)
    End If
    txt_Age.Text = Trim(age1)
End If
            
txt_pridoc.Text = Trim(SEL_TA101_RS("ud_name"))
txt_catEgory.Text = Trim(SEL_TA101_RS("pat_catg"))

SEL_TA101_RS.Close

'# # # # # # # To be done 17/11/2009 # # # # # # # # # #
str = "SELECT * FROM PATOOLSLIB.TA1012 WHERE CASE_NO=?"
Set sel_ta1012_qr = rdb.CreateQuery("", str)

sel_ta1012_qr(0) = Trim(cno)
Set sel_ta1012_rs = sel_ta1012_qr.OpenResultset(1, 3)

If Not sel_ta1012_rs.EOF Then
    txt_provdiag.Text = sel_ta1012_rs("prvdiagdes")
End If
'# # # # # # # # # # # # # # # # # # # # # # # # # # # #

'***********************

l_date = DateAdd("d", -15, Format(DATE11, "DD/MM/YYYY"))
l_date = Format(l_date, "yyyy-mm-dd")
f_date = Year(l_date) & "," & month(l_date) & "," & Day(l_date)

sel_rreq_qr(0) = cno
sel_rreq_qr(1) = f_date
Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)

'***** Lalest 1 *****

'sel_dhemcbc_qr(0) = sel_rreq_rs("reqnno")
'Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 3)

MSFlexGrid1.ColWidth(0) = 1800
MSFlexGrid1.ColAlignment(0) = vbCenter

If Not sel_rreq_rs.EOF Then
    If sel_rreq_rs.RowCount < 3 Then
        
        '+-+-+-+-+ Shraddha +-+-+-+-+
        '+++++
            'Panic value check
            sel_dhem_qr(0) = sel_rreq_rs("REQNNO")
            Set sel_dhem_rs = sel_dhem_qr.OpenResultset(1, 3)
            
            For a = 1 To sel_dhem_rs.RowCount
                If Trim(sel_dhem_rs("PRO_CODE")) = "HAEMOGLOBIN" Then
                    If Trim(sel_dhem_rs("OBSERVALUE")) < 4 Then 'Beep
                        panic = True
                        For S = 1 To 10
                            Beep
                        Next S
                    
                        'MSFlexGrid1.AddItem I & vbTab & sel_chkrreq_rs("CASENO") & vbTab & req_chglen(Trim(sel_chkrreq_rs("reqno"))) & vbTab & sel_chkrreq_rs("TDATE")
                        'I = I + 1
                    End If
                End If
                
                If Trim(sel_dhem_rs("PRO_CODE")) = "PLATELETS" Then
                    If Trim(sel_dhem_rs("OBSERVALUE")) < 5 Then 'Beep
                        panic = True
                        For S = 1 To 10
                            Beep
                        Next S
                    
                        'MSFlexGrid1.AddItem I & vbTab & sel_chkrreq_rs("CASENO") & vbTab & req_chglen(Trim(sel_chkrreq_rs("reqno"))) & vbTab & sel_chkrreq_rs("TDATE")
                        'I = I + 1
                    End If
                End If
                sel_dhem_rs.MoveNext
            Next
            
            '+++++
            
            If panic = True Then
                MSFlexGrid1.Cols = sel_rreq_rs.RowCount + 2
                For c = 2 To sel_rreq_rs.RowCount + 1
                    MSFlexGrid1.ColWidth(c) = 1800
                    'MSFlexGrid1.Height(c)=500
                    MSFlexGrid1.ColAlignment(c) = vbCenter
                    MSFlexGrid1.TextMatrix(0, c) = Trim(sel_rreq_rs("reqnno"))
                    MSFlexGrid1.TextMatrix(1, c) = Trim(sel_rreq_rs("freportdt"))
                    MSFlexGrid1.TextMatrix(2, c) = Trim(sel_rreq_rs("REMARKS"))
                    sel_rreq_rs.MoveNext
                Next c
            Else
'''                MsgBox "No Record(s) found"
'''                'If less than 3 then exit sub
'''                'Set a variable
'''                CBCrec = True
'''                panic = False
'''                Exit Function 'Added on 23/11/09 Shraddha

            '$$$$$$$$$$$$$$$$$$$$
                MSFlexGrid1.Cols = sel_rreq_rs.RowCount + 2
                For c = 2 To sel_rreq_rs.RowCount + 1
                    MSFlexGrid1.ColWidth(c) = 1800
                    'MSFlexGrid1.Height(c)=500
                    MSFlexGrid1.ColAlignment(c) = vbCenter
                    'frm_CBCedit.show
                    MSFlexGrid1.TextMatrix(0, c) = Trim(sel_rreq_rs("reqnno"))
                    MSFlexGrid1.TextMatrix(1, c) = Trim(sel_rreq_rs("freportdt"))
                    MSFlexGrid1.TextMatrix(2, c) = Trim(sel_rreq_rs("REMARKS"))
                    sel_rreq_rs.MoveNext
                Next c
                  
            '$$$$$$$$$$$$$$$$$$$$$$$

            End If
            panic = False
        '+-+-+-+-+-+-+-+-+-+-+-+-+-+-
        
'''        MSFlexGrid1.Cols = sel_rreq_rs.RowCount + 2
'''        For c = 2 To sel_rreq_rs.RowCount + 1
'''            MSFlexGrid1.ColWidth(c) = 1800
'''            'MSFlexGrid1.Height(c)=500
'''            MSFlexGrid1.ColAlignment(c) = vbCenter
'''            'frm_CBCedit.show
'''            MSFlexGrid1.TextMatrix(0, c) = Trim(sel_rreq_rs("reqnno"))
'''            MSFlexGrid1.TextMatrix(1, c) = Trim(sel_rreq_rs("freportdt"))
'''            MSFlexGrid1.TextMatrix(2, c) = Trim(sel_rreq_rs("REMARKS"))
'''            sel_rreq_rs.MoveNext
'''        Next c
    Else
        MSFlexGrid1.Cols = 5
        For c = 2 To 4
            MSFlexGrid1.ColWidth(c) = 1800
            MSFlexGrid1.ColAlignment(c) = vbCenter
            
            MSFlexGrid1.TextMatrix(0, c) = Trim(sel_rreq_rs("reqnno"))
            MSFlexGrid1.TextMatrix(1, c) = Trim(sel_rreq_rs("freportdt"))
            MSFlexGrid1.TextMatrix(2, c) = Trim(sel_rreq_rs("REMARKS"))
            sel_rreq_rs.MoveNext
        Next c
    End If
Else
    'Set a variable
    'CBCrec = True
    '----------------------
    'CHECK FOR PANIC VALUE
    SELHINTFVAL_QR(0) = req_chglen(Trim(R_NO)) 'VAR_REQNO 'Mid(VAR_REQNO, 6, 13)
    Set SELHINTFVAL_RS = SELHINTFVAL_QR.OpenResultset(1, 3)

    If Trim(SELHINTFVAL_RS("HB")) < 4 Or Trim(SELHINTFVAL_RS("PLATELETS")) < 5 Then 'Beep
        'panic_v = True
        For S = 1 To 10
            Beep
        Next S
        frm_CBCedit.show
    End If
'
'
'
'        MSFlexGrid1.Cols = sel_rreq_rs.RowCount + 2
'        For c = 2 To sel_rreq_rs.RowCount + 1
'            MSFlexGrid1.ColWidth(c) = 1800
'            'MSFlexGrid1.Height(c)=500
'            MSFlexGrid1.ColAlignment(c) = vbCenter
'            'frm_CBCedit.show
'            MSFlexGrid1.TextMatrix(0, c) = Trim(sel_rreq_rs("reqnno"))
'            MSFlexGrid1.TextMatrix(1, c) = Trim(sel_rreq_rs("freportdt"))
'            MSFlexGrid1.TextMatrix(2, c) = Trim(sel_rreq_rs("REMARKS"))
'            sel_rreq_rs.MoveNext
'        Next c
'    End If
     '---------------------
End If
'frm_CBCedit.show
'GETDATA_LAST
GRID_DATA

'Call Get_Hinterface(cno)
'********************

End Function

'Public Sub GETDATA_LAST()
'
'Dim num As Integer
'num = Trim(txt_num.Text)
''NUM = NUM - 1
'If opt_patient.Value = True Then
'    MSFlexGrid1.Cols = 1 + num
'
'    SELDHEMLFLAST_QR(0) = Trim(msk_caseno.FormattedText)
'    Set SELDHEMLFLAST_RS = SELDHEMLFLAST_QR.OpenResultset(1, 3)
'
'    MSFlexGrid1.ColWidth(0) = 1800
'    MSFlexGrid1.ColAlignment(0) = vbCenter
'    For c = 1 To num
'        MSFlexGrid1.ColWidth(c) = 1800
'        MSFlexGrid1.ColAlignment(c) = vbCenter
'        MSFlexGrid1.TextMatrix(0, c) = Trim(SELDHEMLFLAST_RS("REQNNO"))
'        MSFlexGrid1.TextMatrix(1, c) = Trim(SELDHEMLFLAST_RS("FREPORTDT"))
'        SELDHEMLFLAST_RS.MoveNext
'    Next c
'
'End If
'End Sub

Public Sub GRID_DATA()
Dim GRIDCOL As Integer
Dim GRIDROW As Integer
r = 0
c = 0
GRIDCOL = MSFlexGrid1.Cols
GRIDROW = MSFlexGrid1.Rows

GRIDCOL = GRIDCOL - 1
GRIDROW = GRIDROW - 1

For c = 2 To GRIDCOL
'For c = GRIDCOL To 1 Step -1
    For r = 3 To GRIDROW
       
        'SELDHEMLF_QR(0) = Trim(MSFlexGrid1.TextMatrix(0, c))
        'SELDHEMLF_QR(1) = Trim(MSFlexGrid1.TextMatrix(r, 0))
        'Set SELDHEMLF_RS = SELDHEMLF_QR.OpenResultset(1, 3)
        If Trim(MSFlexGrid1.TextMatrix(0, c)) <> "" Then
            sel_dhemcbc_qr(0) = Trim(MSFlexGrid1.TextMatrix(0, c))
            sel_dhemcbc_qr(1) = Trim(MSFlexGrid1.TextMatrix(r, 0))
            Set sel_dhemcbc_rs = sel_dhemcbc_qr.OpenResultset(1, 3)
            
            If Not sel_dhemcbc_rs.EOF Then
                If Trim(sel_dhemcbc_rs("OBSERVALUE")) <> "" Then
                    MSFlexGrid1.TextMatrix(r, c) = Trim(sel_dhemcbc_rs("OBSERVALUE"))
                Else
                    MSFlexGrid1.TextMatrix(r, c) = ""
                End If
            End If
        End If
    Next r
Next c
End Sub

Function Get_Hinterface(c_no As String)
Dim col
str = "select * from INTERFACE.HINTERFACE WHERE CASENO=? order by TDATE desc"
Set sel_editHINT_qr = rdb.CreateQuery("", str)

sel_editHINT_qr(0) = Trim(c_no)
Set sel_editHINT_rs = sel_editHINT_qr.OpenResultset(1, 3)

'MSFlexGrid1.Cols = MSFlexGrid1.Cols + 1
'col = MSFlexGrid1.Cols - 1
MSFlexGrid1.Rows = 23

MSFlexGrid1.col = 1
MSFlexGrid1.ColWidth(1) = 1800
For i = 3 To 22
MSFlexGrid1.row = i
   MSFlexGrid1.CellBackColor = &HFFC0C0    'vbGreen
    
Next
'frm_CBCedit.show
If Not sel_editHINT_rs.EOF Then
    
    MSFlexGrid1.ColAlignment(1) = vbCenter
    MSFlexGrid1.TextMatrix(0, 1) = sel_editHINT_rs("reqno") '& "Hinter"
    MSFlexGrid1.TextMatrix(1, 1) = Trim(sel_editHINT_rs("TDATE"))
    MSFlexGrid1.TextMatrix(2, 1) = sel_editHINT_rs("MACHINENO")
    MSFlexGrid1.TextMatrix(3, 1) = sel_editHINT_rs("HB")
    MSFlexGrid1.TextMatrix(4, 1) = sel_editHINT_rs("RBC")
    MSFlexGrid1.TextMatrix(5, 1) = sel_editHINT_rs("HCT")
    MSFlexGrid1.TextMatrix(6, 1) = sel_editHINT_rs("MCV")
    MSFlexGrid1.TextMatrix(7, 1) = sel_editHINT_rs("MCH")
    MSFlexGrid1.TextMatrix(8, 1) = sel_editHINT_rs("MCHC")
    MSFlexGrid1.TextMatrix(9, 1) = sel_editHINT_rs("RDW")
    MSFlexGrid1.TextMatrix(10, 1) = sel_editHINT_rs("PLATELETS")
    MSFlexGrid1.TextMatrix(11, 1) = sel_editHINT_rs("MPV")
    MSFlexGrid1.TextMatrix(12, 1) = sel_editHINT_rs("TLC") '"WBC COUNT"
    MSFlexGrid1.TextMatrix(13, 1) = sel_editHINT_rs("NEUTROPHIL") '"NEUTROPHIL %"
    MSFlexGrid1.TextMatrix(14, 1) = sel_editHINT_rs("LYMPHOCYTE") '"LYMPHOCYTE %"
    MSFlexGrid1.TextMatrix(15, 1) = sel_editHINT_rs("MONOCYTES") '"MONOCYTE %"
    MSFlexGrid1.TextMatrix(16, 1) = sel_editHINT_rs("EOSINOPHIL") '"EOSINOPHIL %"
    MSFlexGrid1.TextMatrix(17, 1) = sel_editHINT_rs("BASOPHILS") '"BASOPHIL %"
    MSFlexGrid1.TextMatrix(18, 1) = "" 'sel_editHINT_rs("NEUTROPHIL") '"NEUTROPHIL ABS"
    MSFlexGrid1.TextMatrix(19, 1) = "" 'sel_editHINT_rs("LYMPHOCYTE") '"LYMPHOCYTE ABS"
    MSFlexGrid1.TextMatrix(20, 1) = "" 'sel_editHINT_rs("MONOCYTES") '"MONOCYTE ABS"
    MSFlexGrid1.TextMatrix(21, 1) = "" 'sel_editHINT_rs("EOSINOPHIL") '"EOSINOPHIL ABS"
    MSFlexGrid1.TextMatrix(22, 1) = "" 'sel_editHINT_rs("BASOPHILS") '"BASOPHIL ABS"
End If
'Next r

End Function

Private Sub MSFlexGrid1_Click()

txt_hinter.Visible = True
If MSFlexGrid1.col = 1 And MSFlexGrid1.row <= 17 Then
txt_hinter.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, MSFlexGrid1.col)

txt_hinter.Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft
txt_hinter.Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
txt_hinter.Width = MSFlexGrid1.CellWidth
txt_hinter.Height = MSFlexGrid1.CellHeight
txt_hinter.SelStart = 0
txt_hinter.SelLength = Len(txt_hinter.Text)
txt_hinter.SetFocus
Cal_Abs_values
End If
Exit Sub
End Sub

Private Sub cmd_commit_Click()

cmd_commit.Caption = "COMMIT"
cmd_commit.BackColor = &H80000003
cmd_commit.Enabled = True
cmd_commit.Visible = True
'Check if record is already inserted if not then only insert else don't

sel_dhem_qr(0) = req_chglen(Trim(R_NO))
Set sel_dhem_rs = sel_dhem_qr.OpenResultset(1, 3)

If sel_dhem_rs.EOF Then
    CBCDETAIL 'INSERTING INTO DHEMCBCDET
End If

sel_dhem_rs.Close

'Cal_Abs_values
UPDATE_Recs 'UPDATING HINTERFACE, RREQUISITN & RTIMELOG
PRINTING
MYDELAY (20)
txt_hinter.Visible = False
Unload Me
'Me.Hide
'frm_CBCworklist.MSFlexGrid1.RemoveItem id
frm_CBCworklist.cmd_refresh_Click
frm_CBCworklist.show
End Sub

Private Sub Cal_Abs_values()
'SELHINTFVAL_QR(0) = req_chglen(Trim(R_NO)) 'Mid(VAR_REQNO, 6, 13)
'Set SELHINTFVAL_RS = SELHINTFVAL_QR.OpenResultset(1, 1)

'NTPHLS = Val(SELHINTFVAL_RS("TLC")) * (Val(SELHINTFVAL_RS("NEUTROPHIL")) / 100)
NTPHLS = (MSFlexGrid1.TextMatrix(12, 1)) * (MSFlexGrid1.TextMatrix(13, 1) / 100)
MSFlexGrid1.TextMatrix(18, 1) = NTPHLS
'LMPCYT = Val(SELHINTFVAL_RS("TLC")) * (Val(SELHINTFVAL_RS("LYMPHOCYTE")) / 100)
LMPCYT = (MSFlexGrid1.TextMatrix(12, 1)) * (MSFlexGrid1.TextMatrix(14, 1) / 100)
MSFlexGrid1.TextMatrix(19, 1) = LMPCYT
'MNCYTE = Val(SELHINTFVAL_RS("TLC")) * (Val(SELHINTFVAL_RS("MONOCYTES")) / 100)
MNCYTE = (MSFlexGrid1.TextMatrix(12, 1)) * (MSFlexGrid1.TextMatrix(15, 1) / 100)
MSFlexGrid1.TextMatrix(20, 1) = MNCYTE
'ESNPHL = Val(SELHINTFVAL_RS("TLC")) * (Val(SELHINTFVAL_RS("EOSINOPHIL")) / 100)
ESNPHL = (MSFlexGrid1.TextMatrix(12, 1)) * (MSFlexGrid1.TextMatrix(16, 1) / 100)
MSFlexGrid1.TextMatrix(21, 1) = ESNPHL
'BSPHLS = Val(SELHINTFVAL_RS("TLC")) * (Val(SELHINTFVAL_RS("BASOPHILS")) / 100)
BSPHLS = (MSFlexGrid1.TextMatrix(12, 1)) * (MSFlexGrid1.TextMatrix(17, 1) / 100)
MSFlexGrid1.TextMatrix(22, 1) = BSPHLS
End Sub

Public Sub CBCDETAIL()

Dim NTPHLS, LMPCYT, MNCYTE, ESNPHL, BSPHLS As Variant

'On Error GoTo MSG:

            SELCASENO_QR(0) = req_chglen(Trim(R_NO)) '"UZZHL" + SELHINTFVAL_RS("REQNO"))
            Set SELCASENO_RS = SELCASENO_QR.OpenResultset(1, 1)
            
            If SELCASENO_RS("PATNSTAT") <> "ST" Then
                SA = SELCASENO_RS("CASENO") 'Trim(MSFlexGrid1.TextMatrix(I, 1))
                SELSEXAGE_QR(0) = Trim(SA)
                Set SELSEXAGE_RS = SELSEXAGE_QR.OpenResultset(1, 1)
                age = DateDiff("YYYY", SELSEXAGE_RS("BIRTH_DATE"), CStr(DATE11))
                SEX = Trim(SELSEXAGE_RS("SEX"))
            Else
                SA = req_chglen(Trim(MSFlexGrid1.TextMatrix(i, 2)))
                SELSTAFF_QR(0) = Trim(SA)
                Set SELSTAFF_RS = SELSTAFF_QR.OpenResultset(1, 1)
                age = SELSTAFF_RS("AGE")
                SEX = SELSTAFF_RS("SEX")
            End If
            
            
            VAR_REQNO = req_chglen(Trim(R_NO))
            SELHINTFVAL_QR(0) = VAR_REQNO 'Mid(VAR_REQNO, 6, 13)
            Set SELHINTFVAL_RS = SELHINTFVAL_QR.OpenResultset(1, 1)
            
            'AGE = DateDiff("YYYY", SELSEXAGE_RS("BIRTH_DATE"), CStr(date11))

            INSDHEMCBCDET_QR(0) = req_chglen(Trim(R_NO)) '"UZZHL" + SELHINTFVAL_RS("REQNO")
        
            If Len(Trim(SELHINTFVAL_RS("HB"))) > 0 Then

                    INSDHEMCBCDET_QR(1) = "HAEMOGLOBIN"
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("HB"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(3, 1)
                    HBAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "HB"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "HB"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE
                    If Not SELDHEMCBCRS.EOF Then
                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    End If
                    If Val(Trim(SELHINTFVAL_RS("HB"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("HB"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("HB"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("HB"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        
                        If Not SELDHEMCBCRS.EOF Then
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        Else
                        INSDHEMCBCDET_QR(7) = ""
                        INSDHEMCBCDET_QR(8) = ""
                        End If
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute


            End If

            If Len(Trim(SELHINTFVAL_RS("RBC"))) > 0 Then

                    INSDHEMCBCDET_QR(1) = "RBC COUNT"
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("RBC"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(4, 1)
                    HBAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "RBC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "RBC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE
                    If Not SELDHEMCBCRS.EOF Then
                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    End If
                    If Val(Trim(SELHINTFVAL_RS("RBC"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("RBC"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("RBC"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("RBC"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        If Not SELDHEMCBCRS.EOF Then
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        Else
                        INSDHEMCBCDET_QR(7) = ""
                        INSDHEMCBCDET_QR(8) = ""
                        End If
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute



            End If

            If Len(Trim(SELHINTFVAL_RS("HCT"))) > 0 Then


                    INSDHEMCBCDET_QR(1) = "HAEMATOCRIT"
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("HCT"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(5, 1)
                    HBAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "HCT"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "HCT"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE
                    If Not SELDHEMCBCRS.EOF Then
                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    End If

                    If Val(Trim(SELHINTFVAL_RS("HCT"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("HCT"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("HCT"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("HCT"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        If Not SELDHEMCBCRS.EOF Then
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        Else
                        INSDHEMCBCDET_QR(7) = ""
                        INSDHEMCBCDET_QR(8) = ""
                        End If
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute



            End If

            If Len(Trim(SELHINTFVAL_RS("MCV"))) > 0 Then


                    INSDHEMCBCDET_QR(1) = "MCV"
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("MCV"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(6, 1)
                    HBAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "MCV"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "MCV"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE
                    If Not SELDHEMCBCRS.EOF Then
                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    End If

                    If Val(Trim(SELHINTFVAL_RS("MCV"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("MCV"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("MCV"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("MCV"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        If Not SELDHEMCBCRS.EOF Then
                            INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                            INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        Else
                            INSDHEMCBCDET_QR(7) = ""
                            INSDHEMCBCDET_QR(8) = ""
                        End If
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute



            End If

            If Len(Trim(SELHINTFVAL_RS("MCH"))) > 0 Then


                    INSDHEMCBCDET_QR(1) = "MCH"
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("MCH"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(7, 1)
                    HBAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "MCH"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "MCH"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE
                    If Not SELDHEMCBCRS.EOF Then
                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    End If

                    If Val(Trim(SELHINTFVAL_RS("MCH"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("MCH"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("MCH"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("MCH"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        If Not SELDHEMCBCRS.EOF Then
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        Else
                        INSDHEMCBCDET_QR(7) = ""
                        INSDHEMCBCDET_QR(8) = ""
                        End If
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute



            End If

            If Len(Trim(SELHINTFVAL_RS("MCHC"))) > 0 Then


                    INSDHEMCBCDET_QR(1) = "MCHC"
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("MCHC"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(8, 1)
                    HBAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "MCHC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "MCHC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE
                    If Not SELDHEMCBCRS.EOF Then
                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    End If

                    If Val(Trim(SELHINTFVAL_RS("MCHC"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("MCHC"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("MCHC"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("MCHC"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        If Not SELDHEMCBCRS.EOF Then
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        Else
                            INSDHEMCBCDET_QR(7) = ""
                            INSDHEMCBCDET_QR(8) = ""
                        End If
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute



            End If

            If Len(Trim(SELHINTFVAL_RS("RDW"))) > 0 Then


                    INSDHEMCBCDET_QR(1) = "RDW"
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("RDW"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(9, 1)
                    'HBAGE
                    RDWAGE
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "RDW"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "RDW"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE
                    If Not SELDHEMCBCRS.EOF Then
                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    End If

                    If Val(Trim(SELHINTFVAL_RS("RDW"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("RDW"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("RDW"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("RDW"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        If Not SELDHEMCBCRS.EOF Then
                            INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                            INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        Else
                            INSDHEMCBCDET_QR(7) = ""
                            INSDHEMCBCDET_QR(8) = ""
                        End If
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute



            End If

            If Len(Trim(SELHINTFVAL_RS("PLATELETS"))) > 0 Then


                    INSDHEMCBCDET_QR(1) = "PLATELETS"
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("PLATELETS"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(10, 1)
                    HBAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Platelets"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Platelets"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE
                    If Not SELDHEMCBCRS.EOF Then
                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))
                    End If
    
                    If Val(Trim(SELHINTFVAL_RS("PLATELETS"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("PLATELETS"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("PLATELETS"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("PLATELETS"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        If Not SELDHEMCBCRS.EOF Then
                            INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                            INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        Else
                            INSDHEMCBCDET_QR(7) = ""
                            INSDHEMCBCDET_QR(8) = ""
                        
                        End If
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute



            End If

            If Len(Trim(SELHINTFVAL_RS("MPV"))) > 0 Then


                    INSDHEMCBCDET_QR(1) = "MPV"
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("MPV"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(11, 1)
                    'HBAGE
                    MPVAGE
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "MPV"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "MPV"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE

                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))


                    If Val(Trim(SELHINTFVAL_RS("MPV"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("MPV"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("MPV"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("MPV"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute


'&H80000013&
'&H00FFC0C0&

            End If

            If Len(Trim(SELHINTFVAL_RS("TLC"))) > 0 Then

                    INSDHEMCBCDET_QR(1) = "WBC COUNT"
                    'Show the value in FlexGrid
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("TLC"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(12, 1)
                    '19/11/09
                    MSFlexGrid1.TextMatrix(12, 1) = Trim(SELHINTFVAL_RS("TLC"))
                    
                    HBAGE
                    'TLCAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "TLC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "TLC"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE

                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    If Val(Trim(SELHINTFVAL_RS("TLC"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("TLC"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("TLC"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("TLC"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute


            End If

            If Len(Trim(SELHINTFVAL_RS("NEUTROPHIL"))) > 0 Then


                    INSDHEMCBCDET_QR(1) = "NEUTROPHIL %"
                   
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("NEUTROPHIL"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(13, 1)
                    
                    HBAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "neutrophil %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "neutrophil %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE

                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    If Val(Trim(SELHINTFVAL_RS("NEUTROPHIL"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("NEUTROPHIL"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("NEUTROPHIL"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("NEUTROPHIL"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute



            End If

            If Len(Trim(SELHINTFVAL_RS("LYMPHOCYTE"))) > 0 Then


                    INSDHEMCBCDET_QR(1) = "LYMPHOCYTE %"
                    
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("LYMPHOCYTE"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(14, 1)
                    LYMPHOCYTEAGE
                    'HBAGE
                    
                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Lymphocyte %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Lymphocyte %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE

                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    If Val(Trim(SELHINTFVAL_RS("LYMPHOCYTE"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("LYMPHOCYTE"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("LYMPHOCYTE"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("LYMPHOCYTE"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute



            End If

            If Len(Trim(SELHINTFVAL_RS("MONOCYTES"))) > 0 Then


                    INSDHEMCBCDET_QR(1) = "MONOCYTE %"
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("MONOCYTES"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(15, 1)
                    'HBAGE
                    MONOCYTEAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Monocyte %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Monocyte %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE

                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    If Val(Trim(SELHINTFVAL_RS("MONOCYTES"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("MONOCYTES"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("MONOCYTES"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("MONOCYTES"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute



            End If

            If Len(Trim(SELHINTFVAL_RS("EOSINOPHIL"))) > 0 Then


                    INSDHEMCBCDET_QR(1) = "EOSINOPHIL %"
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("EOSINOPHIL"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(16, 1)
                    HBAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "eosinophils %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "eosinophils %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE

                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    If Val(Trim(SELHINTFVAL_RS("EOSINOPHIL"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("EOSINOPHIL"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("EOSINOPHIL"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("EOSINOPHIL"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute



            End If

            If Len(Trim(SELHINTFVAL_RS("BASOPHILS"))) > 0 Then


                    INSDHEMCBCDET_QR(1) = "BASOPHIL %"
                    'INSDHEMCBCDET_QR(2) = Trim(SELHINTFVAL_RS("BASOPHILS"))
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(17, 1)
                    'HBAGE
                    BASOPHILSAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Basophils %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Basophils %"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE

                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    If Val(Trim(SELHINTFVAL_RS("BASOPHILS"))) <= Val(HBHV) And Val(Trim(SELHINTFVAL_RS("BASOPHILS"))) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(SELHINTFVAL_RS("BASOPHILS"))) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(SELHINTFVAL_RS("BASOPHILS"))) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute



            End If

            If NTPHLABSFLAG = False Then


                    INSDHEMCBCDET_QR(1) = "NEUTROPHIL ABS"
                    'Show the value in FlexGrid
                    'NTPHLS = Val(SELHINTFVAL_RS("TLC")) * (SELHINTFVAL_RS("NEUTROPHIL") / 100)
                    NTPHLS = Val(SELHINTFVAL_RS("TLC")) * (Val(SELHINTFVAL_RS("NEUTROPHIL")) / 100)
                    'INSDHEMCBCDET_QR(2) = NTPHLS
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(18, 1)
                    MSFlexGrid1.TextMatrix(18, 1) = NTPHLS
                    HBAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Neutrophils abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Neutrophils abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE

                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    If Val(Trim(NTPHLS)) <= Val(HBHV) And Val(Trim(NTPHLS)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(NTPHLS)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(NTPHLS)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute

                    'NTPHLABSFLAG = True

            End If


            If LMPCYTEABSFLAG = False Then


                    INSDHEMCBCDET_QR(1) = "LYMPHOCYTE ABS"
                    'Show the value in FlexGrid
                    LMPCYT = Val(SELHINTFVAL_RS("TLC")) * (Val(SELHINTFVAL_RS("LYMPHOCYTE")) / 100)
                    'INSDHEMCBCDET_QR(2) = LMPCYT
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(19, 1)
                    
                    MSFlexGrid1.TextMatrix(19, 1) = LMPCYT
                    
                    HBAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Lymphocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Lymphocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE

                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    If Val(Trim(LMPCYT)) <= Val(HBHV) And Val(Trim(LMPCYT)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(LMPCYT)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(LMPCYT)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute

                    'LMPCYTEABSFLAG = True

            End If


            If MNCYTEABSFLAG = False Then


                    INSDHEMCBCDET_QR(1) = "MONOCYTE ABS"
                    'Show the value in FlexGrid
                    MNCYTE = Val(SELHINTFVAL_RS("TLC")) * (Val(SELHINTFVAL_RS("MONOCYTES")) / 100)
                    'INSDHEMCBCDET_QR(2) = MNCYTE
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(20, 1)
                    
                    MSFlexGrid1.TextMatrix(20, 1) = MNCYTE
                    
                    HBAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Monocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Monocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE

                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    If Val(Trim(MNCYTE)) <= Val(HBHV) And Val(Trim(MNCYTE)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(MNCYTE)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(MNCYTE)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute

                    'MNCYTEABSFLAG = True

            End If



            If ESNPHLABSFLAG = False Then


                    INSDHEMCBCDET_QR(1) = "EOSINOPHIL ABS"
                    'Show the value in FlexGrid
                    ESNPHL = Val(SELHINTFVAL_RS("TLC")) * (Val(SELHINTFVAL_RS("EOSINOPHIL")) / 100)
                    'INSDHEMCBCDET_QR(2) = ESNPHL
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(21, 1)
                    MSFlexGrid1.TextMatrix(21, 1) = ESNPHL
                    
                    HBAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Monocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Monocyte abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE

                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    If Val(Trim(ESNPHL)) <= Val(HBHV) And Val(Trim(ESNPHL)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(ESNPHL)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(ESNPHL)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute

                    'ESNPHLABSFLAG = True
            End If



            If BSPHLABSFLAG = False Then


                    INSDHEMCBCDET_QR(1) = "BASOPHIL ABS"
                    'Show the value in FlexGrid
                    BSPHLS = Val(SELHINTFVAL_RS("TLC")) * (Val(SELHINTFVAL_RS("BASOPHILS")) / 100)
                    'INSDHEMCBCDET_QR(2) = BSPHLS
                    INSDHEMCBCDET_QR(2) = MSFlexGrid1.TextMatrix(22, 1)
                    MSFlexGrid1.TextMatrix(22, 1) = BSPHLS
                    
                    'HBAGE
                    BASOPHILSAGE

                    If SELDHEMCBCQRFLAG = False Then
                            SELDHEMCBCQR(0) = "Basophils abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            Set SELDHEMCBCRS = SELDHEMCBCQR.OpenResultset(1, 3)
                            SELDBCALKPHSFLAG = True
                    Else
                            SELDHEMCBCQR(0) = "Basophils abs"
                            SELDHEMCBCQR(1) = HAGE
                            SELDHEMCBCQR(2) = HBSEX
                            SELDHEMCBCRS.MoreResults
                            SELDHEMCBCRS.Requery
                    End If

                    'RANGE = SELDHEMCBCRS!RANGE

                    I1 = InStr(1, Trim(SELDHEMCBCRS("RANGE")), "-")
                    HBLV = Mid(Trim(SELDHEMCBCRS("RANGE")), 1, I1 - 1)
                    HBHV = Mid(Trim(SELDHEMCBCRS("RANGE")), I1 + 1, Len(Trim(SELDHEMCBCRS("RANGE"))))

                    If Val(Trim(BSPHLS)) <= Val(HBHV) And Val(Trim(BSPHLS)) >= Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "N"
                    ElseIf Val(Trim(BSPHLS)) < Val(HBLV) Then
                        INSDHEMCBCDET_QR(3) = "L"
                    ElseIf Val(Trim(BSPHLS)) > Val(HBHV) Then
                        INSDHEMCBCDET_QR(3) = "H"
                    End If

                        INSDHEMCBCDET_QR(4) = ""
                        INSDHEMCBCDET_QR(5) = ""
                        INSDHEMCBCDET_QR(6) = SELHINTFVAL_RS("MACHINENO")
                        INSDHEMCBCDET_QR(7) = SELDHEMCBCRS("UOM")
                        INSDHEMCBCDET_QR(8) = SELDHEMCBCRS("RANGE")
                        INSDHEMCBCDET_QR(9) = ""

                        INSDHEMCBCDET_QR.Execute

                    'BSPHLABSFLAG = True
            End If
            
            SAVEERROR = False
            
            Exit Sub

MSG:
    INSCBCINTFERR_QR(0) = VAR_REQNO 'Trim(MSFlexGrid1.TextMatrix(I, 2))
    INSCBCINTFERR_QR(1) = Err.Number
    INSCBCINTFERR_QR(2) = "SAVE"
    INSCBCINTFERR_QR(3) = Err.Description
    INSCBCINTFERR_QR(4) = Date
    INSCBCINTFERR_QR(5) = Time
    INSCBCINTFERR_QR(6) = "N"
    INSCBCINTFERR_QR.Execute
    SAVEERROR = True
    'ERROR = True
    
        
End Sub


Public Sub UPDATE_Recs()

On Error GoTo MSG:

    Dim rno As String

    rno = req_chglen(Trim(R_NO)) 'req_chglen(Trim(MSFlexGrid1.TextMatrix(I, 2)))
    'R_NO
    
    updhinterface_qr(0) = rno  'selhintfrs("REQNO")
    updhinterface_qr.Execute

    updrrequisitn1_qr(0) = Trim(TXT_OPERATOR.Text)
    updrrequisitn1_qr(1) = Trim(msk_sysdate.FormattedText)
    updrrequisitn1_qr(2) = "F"
    updrrequisitn1_qr(3) = "Y"
    updrrequisitn1_qr(4) = Trim(rno) 'req_chglen(Trim(MSFlexGrid1.TextMatrix(I, 2))) '"UZZHL" + selhintfrs("REQNO")
    updrrequisitn1_qr.Execute

    UPDRTIMEC_QR(0) = Trim(msk_sysdate.FormattedText)
    UPDRTIMEC_QR(1) = Trim(msk_systime.FormattedText)
    UPDRTIMEC_QR(2) = Trim(msk_sysdate.FormattedText)
    UPDRTIMEC_QR(3) = Trim(msk_systime.FormattedText)
    UPDRTIMEC_QR(4) = Trim(rno) 'req_chglen(Trim(MSFlexGrid1.TextMatrix(I, 2))) '"UZZHL" + selhintfrs("REQNO")
    UPDRTIMEC_QR.Execute

    Exit Sub

MSG:

    INSCBCINTFERR_QR(0) = req_chglen(Trim(R_NO))
    INSCBCINTFERR_QR(1) = Err.Number
    INSCBCINTFERR_QR(2) = "COMMIT"
    INSCBCINTFERR_QR(3) = Err.Description
    INSCBCINTFERR_QR(4) = Date
    INSCBCINTFERR_QR(5) = Time
    INSCBCINTFERR_QR(6) = "N"
    INSCBCINTFERR_QR.Execute
    Error = True

End Sub

Public Sub PRINTING()

On Error GoTo MSG:

    Dim str1 As String
    Dim str2 As String
    
    DoEvents
    
    sel_ta1012_qr(0) = Trim(CASENO) 'Trim(MSFlexGrid1.TextMatrix(I, 1))
    Set sel_ta1012_rs = sel_ta1012_qr.OpenResultset(1, 1)
    
    str1 = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Time
              
    'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(Trim(MSFlexGrid1.TextMatrix(I, 2))) & "'"
     CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & req_chglen(Trim(R_NO)) & "'"
             
    SELPATNSTAT_QR(0) = req_chglen(Trim(R_NO))
    Set SELPATNSTAT_RS = SELPATNSTAT_QR.OpenResultset(1, 1)
            
    If SELPATNSTAT_RS("PATNSTAT") <> "ST" Then
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISAUTOCBC.RPT"
        CrystalReport1.Formulas(2) = "PROVDIA = '" & sel_ta1012_rs("PRVDIAGDES") & "'"
    Else
        CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISAUTOCBCS.RPT"
    End If
            
            CrystalReport1.Formulas(0) = "REQNO = '" & MSFlexGrid1.TextMatrix(i, 2) & "'"
            
            
            CrystalReport1.Formulas(5) = "PROFIN = '" & "AUTOMATED COMPLETE BLOOD COUNT REPORT" & "'"

'            STR2 = ""
'            STR2 = "Operated by " & Trim(TXT_OPERATOR.Text)
'
'                If Len(Trim(TXT_OPERATOR.Text)) > 0 Then
'                    CrystalReport1.Formulas(9) = "ENT = '" & STR2 & "'"
'                Else
'                     CrystalReport1.Formulas(9) = "ENT = '" & "" & "'"
'                End If
             
            CrystalReport1.Formulas(12) = "USRID = '" & str1 & "'"
            CrystalReport1.WindowState = crptMaximized
            'Always keep to Printer
            'CrystalReport1.Destination = crptToPrinter
            CrystalReport1.Destination = crptToWindow
            CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
            CrystalReport1.Action = 1
               
            Exit Sub
            
MSG:

    INSCBCINTFERR_QR(0) = req_chglen(Trim(MSFlexGrid1.TextMatrix(i, 2)))
    INSCBCINTFERR_QR(1) = Err.Number
    INSCBCINTFERR_QR(2) = "PRINT"
    INSCBCINTFERR_QR(3) = Err.Description
    INSCBCINTFERR_QR(4) = Date
    INSCBCINTFERR_QR(5) = Time
    INSCBCINTFERR_QR(6) = "N"
    INSCBCINTFERR_QR.Execute
    Error = True
    

End Sub
Public Sub HBAGE()


If (age <= 0.2 And age >= "0.00") And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 0.2
    HBSEX = "M"
End If

If (age <= 0.2 And age >= "0.00") And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 0.2
    HBSEX = "F"
End If

If (age <= 0.1 And age >= 0.03) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 0.1
    HBSEX = "M"
End If

If (age <= 0.1 And age >= 0.03) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 0.1
    HBSEX = "F"
End If

If (age <= 0.2 And age > 0.1) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 0.2
    HBSEX = "M"
End If

If (age <= 0.2 And age > 0.1) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 0.2
    HBSEX = "F"
End If

If (age <= 0.3 And age > 0.2) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 0.3
    HBSEX = "M"
End If

If (age <= 0.3 And age > 0.2) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 0.3
    HBSEX = "F"
End If

If (age <= 0.6 And age > 0.3) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 0.6
    HBSEX = "M"
End If

If (age <= 0.6 And age > 0.3) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 0.6
    HBSEX = "F"
End If

If (age <= 2 And age > 0.6) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 2
    HBSEX = "M"
End If

If (age <= 2 And age > 0.6) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 2
    HBSEX = "F"
End If

If (age <= 6 And age > 2) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 6
    HBSEX = "M"
End If

If (age <= 6 And age > 2) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 6
    HBSEX = "F"
End If

If (age <= 12 And age > 6) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 12
    HBSEX = "M"
End If

If (age <= 12 And age > 6) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 12
    HBSEX = "F"
End If

If (age <= 106 And age > 12) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 106
    HBSEX = "M"
End If

If (age <= 106 And age > 12) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 106
    HBSEX = "F"
End If

End Sub
Public Sub RDWAGE()

If (age <= 106 And age >= 0) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 106
    HBSEX = "M"
End If

If (age <= 106 And age >= 0) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 106
    HBSEX = "F"
End If

End Sub
Public Sub MPVAGE()

If (age <= 106 And age >= 0) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 106
    HBSEX = "M"
End If

If (age <= 106 And age >= 0) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 106
    HBSEX = "F"
End If

End Sub
Public Sub LYMPHOCYTEAGE()

If (age <= 0.02 And age >= "0.00") And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 0.02
    HBSEX = "M"
End If

If (age <= 0.02 And age >= "0.00") And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 0.02
    HBSEX = "F"
End If

If (age <= 0.1 And age >= 0.03) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 0.1
    HBSEX = "M"
End If

If (age <= 0.1 And age >= 0.03) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 0.1
    HBSEX = "F"
End If

If (age <= 0.2 And age > 0.1) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 0.2
    HBSEX = "M"
End If

If (age <= 0.2 And age > 0.1) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 0.2
    HBSEX = "F"
End If

If (age <= 0.6 And age > 0.2) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 0.6
    HBSEX = "M"
End If

If (age <= 0.6 And age > 0.2) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 0.6
    HBSEX = "F"
End If

If (age <= 6 And age > 0.6) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 6
    HBSEX = "M"
End If

If (age <= 6 And age > 0.6) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 6
    HBSEX = "F"
End If

If (age <= 12 And age > 6) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 12
    HBSEX = "M"
End If

If (age <= 12 And age > 6) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 12
    HBSEX = "F"
End If

If (age <= 106 And age > 12) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 106
    HBSEX = "M"
End If

If (age <= 106 And age > 12) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 106
    HBSEX = "F"
End If

End Sub

Public Sub MONOCYTEAGE()

If (age <= 0.2 And age >= "0.00") And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 0.2
    HBSEX = "M"
End If

If (age <= 0.2 And age >= "0.00") And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 0.2
    HBSEX = "F"
End If

If (age <= 0.1 And age >= 0.03) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 0.1
    HBSEX = "M"
End If

If (age <= 0.1 And age >= 0.03) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 0.1
    HBSEX = "F"
End If

If (age <= 0.3 And age > 0.1) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 0.3
    HBSEX = "M"
End If

If (age <= 0.3 And age > 0.1) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 0.3
    HBSEX = "F"
End If

If (age <= 0.6 And age > 0.3) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 0.6
    HBSEX = "M"
End If

If (age <= 0.6 And age > 0.3) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 0.6
    HBSEX = "F"
End If

If (age <= 2 And age > 0.6) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 2
    HBSEX = "M"
End If

If (age <= 2 And age > 0.6) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 2
    HBSEX = "F"
End If

If (age <= 6 And age > 2) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 6
    HBSEX = "M"
End If

If (age <= 6 And age > 2) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 6
    HBSEX = "F"
End If

If (age <= 12 And age > 6) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 12
    HBSEX = "M"
End If

If (age <= 12 And age > 6) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 12
    HBSEX = "F"
End If

If (age <= 106 And age > 12) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 106
    HBSEX = "M"
End If

If (age <= 106 And age > 12) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 106
    HBSEX = "F"
End If

End Sub


Public Sub BASOPHILSAGE()

If (age <= 1 And age >= "0.00") And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 1
    HBSEX = "M"
End If

If (age <= 1 And age >= "0.00") And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 1
    HBSEX = "F"
End If

If (age <= 4 And age > 1) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 4
    HBSEX = "M"
End If

If (age <= 4 And age > 1) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 4
    HBSEX = "F"
End If

If (age <= 6 And age > 4) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 6
    HBSEX = "M"
End If

If (age <= 6 And age > 4) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 6
    HBSEX = "F"
End If

If (age <= 10 And age > 6) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 10
    HBSEX = "M"
End If

If (age <= 10 And age > 6) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 10
    HBSEX = "F"
End If

If (age <= 21 And age > 10) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 21
    HBSEX = "M"
End If

If (age <= 21 And age > 10) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 21
    HBSEX = "F"
End If

If (age <= 106 And age > 21) And SEX = "M" Then 'And Trim(SELSEXAGE_RS("SEX")) = "M" Then
    HAGE = 106
    HBSEX = "M"
End If

If (age <= 106 And age > 21) And SEX = "F" Then 'And Trim(SELSEXAGE_RS("SEX")) = "F" Then
    HAGE = 106
    HBSEX = "F"
End If

End Sub

Private Sub txt_hinter_KeyPress(KeyAscii As Integer)

MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1) = txt_hinter.Text
If KeyAscii = 13 And MSFlexGrid1.row < 17 Then
    MSFlexGrid1.row = MSFlexGrid1.row + 1
    txt_hinter.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    txt_hinter.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, MSFlexGrid1.col)
    Cal_Abs_values
End If
End Sub

Public Sub MYDELAY(inttime As String)

'TXT_MESSAGE = "  Waiting for " & inttime + " Seconds"
Dim sttime
Dim endtime1 As Long
Dim TESTTIME As Long

endtime1 = Timer + inttime
TESTTIME = Timer

Do While Timer < endtime1
    DoEvents
Loop

End Sub
