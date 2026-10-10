VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_TMregister 
   Caption         =   "Form2"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   LinkTopic       =   "Form2"
   ScaleHeight     =   9390
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   1800
      TabIndex        =   2
      Top             =   1800
      Width           =   8295
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         Height          =   855
         Left            =   2640
         TabIndex        =   12
         Top             =   1440
         Width           =   5175
         Begin MSMask.MaskEdBox msk_fromdt 
            Height          =   375
            Left            =   1320
            TabIndex        =   13
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   -2147483626
            ForeColor       =   4210752
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
         Begin MSMask.MaskEdBox msk_todt 
            Height          =   375
            Left            =   3720
            TabIndex        =   14
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   -2147483626
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0FF&
            Caption         =   "To Date"
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
            Left            =   2880
            TabIndex        =   16
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0FF&
            Caption         =   "From Date"
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
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton CMD_CANCEL 
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
         Height          =   495
         Left            =   2760
         TabIndex        =   11
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton CMDREPORT 
         Caption         =   "Show &Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   10
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton CMD_BACK 
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
         Height          =   495
         Left            =   6480
         TabIndex        =   9
         Top             =   3000
         Width           =   1215
      End
      Begin VB.ComboBox cmb_mod 
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
         ItemData        =   "TM_register.frx":0000
         Left            =   4680
         List            =   "TM_register.frx":0002
         TabIndex        =   8
         Top             =   600
         Width           =   2415
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0FF&
         Height          =   2415
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
         Begin VB.OptionButton OPT_1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Drug Register"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton OPT_2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Tumor Marker Investication Register"
            Height          =   615
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton OPT_3 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Electro Phoresis Register"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Shape Shape2 
            BorderStyle     =   6  'Inside Solid
            BorderWidth     =   2
            Height          =   2295
            Left            =   0
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.TextBox TXT_MESSAGE 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   4320
         Width           =   5655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Select Modality :"
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
         Left            =   2640
         TabIndex        =   18
         Top             =   600
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         Height          =   5055
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   8295
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "MESSAGE"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   4320
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
      Begin MSMask.MaskEdBox Msk_caseno 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
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
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   840
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CR 
      Left            =   360
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "REPORTS OF BIOCHEMICAL INVESTIGATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   840
      Width           =   6975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   480
      TabIndex        =   19
      Top             =   240
      Width           =   10695
   End
End
Attribute VB_Name = "frm_TMregister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SELDM_QR As rdoQuery
Dim SELDM_RS As rdoResultset
Dim SELDMFLAG As Boolean
Dim DEL_QR As rdoQuery
Dim DEL_RS As rdoResultset
Dim DEL_FLG As Boolean
Dim SELRREQ_QR As rdoQuery
Dim SELRREQ_RS As rdoResultset
Dim SELRREQFLAG As Boolean
Dim SELRREQ_QR1 As rdoQuery
Dim SELRREQ_RS1 As rdoResultset

'SHWETA 06-JULEY-2010
Dim SELEL_QR As rdoQuery
Dim SELEL_RS As rdoResultset

Dim SELELREQ_QR As rdoQuery
Dim SELELREQ_RS As rdoResultset

Dim UPDEL_QR As rdoQuery
Dim UPDEL_RS As rdoResultset

Dim DELELTEMP_QR As rdoQuery

'-------------------------



Dim SELRREQFLAG1 As Boolean

Dim selreq_qr As rdoQuery
Dim selreq_rs As rdoResultset
Dim SELREQ_FLAG As Boolean
Dim UPDBIO_QR As rdoQuery
Dim UPDBIO_RS As rdoResultset

Dim UPDBIO_FLAG As Boolean
Dim RS As rdoResultset
Dim INSBIOTEMP_QR As rdoQuery
Dim INSBIOTEMP_RS As rdoResultset
''
Dim SELTM_QR As rdoQuery
Dim SELTM_RS As rdoResultset
Dim updtumtemp_qr As rdoQuery
Dim DELTUMTEMP_QR As rdoQuery
Dim seltumtemp_qr As rdoQuery
Dim SELTM_QR1 As rdoQuery
Dim SELTM_RS1 As rdoResultset

Dim INSDRUGREG_QR As rdoQuery
Dim SELDRUGREG_QR As rdoQuery
Dim SELDRUGREG_RS As rdoResultset
Dim UPDDRUGREG_QR As rdoQuery
Dim seldmulproce_qr As rdoQuery
Dim seldmulproce_rs As rdoResultset
Dim DELDRUGREG_QR As rdoQuery
Dim SELDOSE_QR As rdoQuery
Dim SELDOSE_RS As rdoResultset




Private Sub cmb_mod_Click()
'Option1.Enabled = True
'Option2.Enabled = True

Dim mod2 As String
mod2 = Right(Trim(cmb_mod.Text), 2)

If mod2 = "TM" Then
    Frame4.Enabled = True
Else
    OPT_1.Value = False
    OPT_2.Value = False
    OPT_3.Value = False
    
    Frame4.Enabled = False
    
End If

msk_fromdt.SetFocus

End Sub

Private Sub cmd_back_Click()
    msk_fromdt.Mask = ""
    msk_fromdt.Text = ""
    msk_fromdt.Mask = "##/##/####"
    msk_todt.Mask = ""
    msk_todt.Text = ""
    msk_todt.Mask = "##/##/####"
    DIS_MAIN.show
    Me.Hide
    'Unload Me
    DEL_QR.Execute
    DELTUMTEMP_QR.Execute
    DELELTEMP_QR.Execute
    CMDREPORT.Enabled = True
End Sub

Private Sub cmd_cancel_Click()
msk_fromdt.Text = "__/__/____"
msk_todt.Text = "__/__/____"

DEL_QR.Execute
DELTUMTEMP_QR.Execute
DELELTEMP_QR.Execute
'Msk_caseno.Mask = "??>/#####"
'Msk_caseno.Text = ""
cmd_cancel.Enabled = False
CMDREPORT.Enabled = True
End Sub
Private Sub CMDREPORT_Click()
Me.MousePointer = 11
DoEvents



   TXT_MESSAGE.Text = "  DISPLAYING REPORT PLEASE WAIT......"
    MYDELAY (5)


Dim mod1 As String
mod1 = Right(Trim(cmb_mod.Text), 2)


Dim var_date_f As String
Dim var_date_t As String

var_date_f = Year(msk_fromdt.FormattedText) & "," & month(msk_fromdt.FormattedText) & "," & Day(msk_fromdt.FormattedText)
var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)


'' for Biochemistry

If mod1 = "BC" Then

'If Option1.Value = True Then
    SELRREQ_QR1(0) = GET_REVDATE(msk_fromdt.Text)
    SELRREQ_QR1(1) = GET_REVDATE(msk_todt.Text)
    SELRREQ_QR1.Execute
    Set selreq_rs = selreq_qr.OpenResultset(1, 3)
    Do Until selreq_rs.EOF
    DoEvents
    i = selreq_rs!REQNO
    SELDM_QR(0) = i
    SELDM_QR.Execute
    Set SELDM_RS = SELDM_QR.OpenResultset(1, 3)
    UPDBIO_QR(0) = ""
    UPDBIO_QR(1) = ""
    UPDBIO_QR(2) = ""
    UPDBIO_QR(3) = ""
    UPDBIO_QR(4) = ""
    UPDBIO_QR(5) = ""
    UPDBIO_QR(6) = ""
    UPDBIO_QR(7) = ""
    UPDBIO_QR(8) = ""
    UPDBIO_QR(9) = ""
    UPDBIO_QR(10) = ""
    UPDBIO_QR(11) = ""
    UPDBIO_QR(12) = ""
    UPDBIO_QR(13) = ""
    UPDBIO_QR(14) = ""
    UPDBIO_QR(15) = ""
    UPDBIO_QR(16) = ""
    UPDBIO_QR(17) = ""
    UPDBIO_QR(18) = ""
    UPDBIO_QR(19) = ""
    UPDBIO_QR(20) = ""
    UPDBIO_QR(21) = ""
    UPDBIO_QR(22) = ""
    UPDBIO_QR(23) = ""
    UPDBIO_QR(24) = ""
    UPDBIO_QR(25) = ""
    UPDBIO_QR(26) = ""
    UPDBIO_QR(27) = ""
    UPDBIO_QR(28) = ""
    UPDBIO_QR(29) = ""
    UPDBIO_QR(30) = ""
    UPDBIO_QR(31) = ""
    UPDBIO_QR(32) = ""
    UPDBIO_QR(33) = ""
    UPDBIO_QR(34) = ""
    UPDBIO_QR(35) = ""
    UPDBIO_QR(36) = ""
    
    Do Until SELDM_RS.EOF
        If SELDM_RS!PRO_CODE = "F842" Then
            UPDBIO_QR(0) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F843" Then
            UPDBIO_QR(1) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F841" Then
            UPDBIO_QR(2) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F846" Then
            UPDBIO_QR(3) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F847" Then
            UPDBIO_QR(4) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F855" Then
            UPDBIO_QR(5) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F856" Then
            UPDBIO_QR(6) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F857" Then
            UPDBIO_QR(7) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F861" Then
            UPDBIO_QR(8) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F862" Then
            UPDBIO_QR(9) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F863" Then
            UPDBIO_QR(10) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F864" Then
            UPDBIO_QR(11) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F888" Then
            UPDBIO_QR(12) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F890" Then
            UPDBIO_QR(13) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F866" Then
            UPDBIO_QR(14) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F867" Then
            UPDBIO_QR(15) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F868" Then
            UPDBIO_QR(16) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F869" Then
            UPDBIO_QR(17) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F870" Then
            UPDBIO_QR(18) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F871" Then
            UPDBIO_QR(19) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F872" Then
            UPDBIO_QR(20) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F873" Then
            UPDBIO_QR(21) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F874" Then
            UPDBIO_QR(22) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F876" Then
            UPDBIO_QR(23) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F895" Then
            UPDBIO_QR(24) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F896" Then
            UPDBIO_QR(25) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F850" Then
            UPDBIO_QR(26) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F853" Then
            UPDBIO_QR(27) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F851" Then
            UPDBIO_QR(28) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F852" Then
            UPDBIO_QR(29) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F881" Then
            UPDBIO_QR(30) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F882" Then
            UPDBIO_QR(31) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F891" Then
            UPDBIO_QR(32) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F924" Then
            UPDBIO_QR(33) = Trim(SELDM_RS!OBSERVALUE)
        
        ElseIf SELDM_RS!PRO_CODE = "F894" Then
            UPDBIO_QR(34) = Trim(SELDM_RS!OBSERVALUE)
        
        ElseIf SELDM_RS!PRO_CODE = "F893" Then
            UPDBIO_QR(35) = Trim(SELDM_RS!OBSERVALUE)
        
        ElseIf SELDM_RS!PRO_CODE = "F845" Then
            UPDBIO_QR(36) = Trim(SELDM_RS!OBSERVALUE)
        End If
        
        SELDM_RS.MoveNext
     Loop
    
    UPDBIO_QR(37) = Trim(i)
    UPDBIO_QR.Execute
    selreq_rs.MoveNext
Loop
CR.Formulas(0) = "FRMDT= '" & msk_fromdt.Text & "'"
CR.Formulas(1) = "TODT='" & msk_todt.Text & "'"
'CR.ReportFileName = "C:\DISREPORTSEC\BIOCHEMICAL2.rpt"
CR.ReportFileName = "C:\DISREPORTSEC\BIOCHEMICALFinalreport.rpt"
CR.Action = 1
'End If

'If Option2.Value = True Then
'    SELRREQ_QR(0) = Trim(Msk_caseno.Text)
'    SELRREQ_QR.Execute
'    Set selreq_rs = selreq_qr.OpenResultset(1, 3)
'    Do Until selreq_rs.EOF
'    i = selreq_rs!reqno
'    SELDM_QR(0) = i
'    SELDM_QR.Execute
'    Set SELDM_RS = SELDM_QR.OpenResultset(1, 3)
'    UPDBIO_QR(0) = ""
'    UPDBIO_QR(1) = ""
'    UPDBIO_QR(2) = ""
'    UPDBIO_QR(3) = ""
'    UPDBIO_QR(4) = ""
'    UPDBIO_QR(5) = ""
'    UPDBIO_QR(6) = ""
'    UPDBIO_QR(7) = ""
'    UPDBIO_QR(8) = ""
'    UPDBIO_QR(9) = ""
'    UPDBIO_QR(10) = ""
'    UPDBIO_QR(11) = ""
'    UPDBIO_QR(12) = ""
'    UPDBIO_QR(13) = ""
'    UPDBIO_QR(14) = ""
'    UPDBIO_QR(15) = ""
'    UPDBIO_QR(16) = ""
'    UPDBIO_QR(17) = ""
'    UPDBIO_QR(18) = ""
'    UPDBIO_QR(19) = ""
'    UPDBIO_QR(20) = ""
'    UPDBIO_QR(21) = ""
'    UPDBIO_QR(22) = ""
'    UPDBIO_QR(23) = ""
'    UPDBIO_QR(24) = ""
'    UPDBIO_QR(25) = ""
'    UPDBIO_QR(26) = ""
'    UPDBIO_QR(27) = ""
'    UPDBIO_QR(28) = ""
'    UPDBIO_QR(29) = ""
'    UPDBIO_QR(30) = ""
'    UPDBIO_QR(31) = ""
'    UPDBIO_QR(32) = ""
'    UPDBIO_QR(33) = ""
'    UPDBIO_QR(34) = ""
'    UPDBIO_QR(35) = ""
'    UPDBIO_QR(36) = ""
'
'    Do Until SELDM_RS.EOF
'        If SELDM_RS!PRO_CODE = "F842" Then
'            UPDBIO_QR(0) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F843" Then
'            UPDBIO_QR(1) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F841" Then
'            UPDBIO_QR(2) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F846" Then
'            UPDBIO_QR(3) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F847" Then
'            UPDBIO_QR(4) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F855" Then
'            UPDBIO_QR(5) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F856" Then
'            UPDBIO_QR(6) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F857" Then
'            UPDBIO_QR(7) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F861" Then
'            UPDBIO_QR(8) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F862" Then
'            UPDBIO_QR(9) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F863" Then
'            UPDBIO_QR(10) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F864" Then
'            UPDBIO_QR(11) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F888" Then
'            UPDBIO_QR(12) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F890" Then
'            UPDBIO_QR(13) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F866" Then
'            UPDBIO_QR(14) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F867" Then
'            UPDBIO_QR(15) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F868" Then
'            UPDBIO_QR(16) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F869" Then
'            UPDBIO_QR(17) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F870" Then
'            UPDBIO_QR(18) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F871" Then
'            UPDBIO_QR(19) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F872" Then
'            UPDBIO_QR(20) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F873" Then
'            UPDBIO_QR(21) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F874" Then
'            UPDBIO_QR(22) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F876" Then
'            UPDBIO_QR(23) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F895" Then
'            UPDBIO_QR(24) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F896" Then
'            UPDBIO_QR(25) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F850" Then
'            UPDBIO_QR(26) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F853" Then
'            UPDBIO_QR(27) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F851" Then
'            UPDBIO_QR(28) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F852" Then
'            UPDBIO_QR(29) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F881" Then
'            UPDBIO_QR(30) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F882" Then
'            UPDBIO_QR(31) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F891" Then
'            UPDBIO_QR(32) = Trim(SELDM_RS!OBSERVALUE)
'
'        ElseIf SELDM_RS!PRO_CODE = "F924" Then
'            UPDBIO_QR(33) = Trim(SELDM_RS!OBSERVALUE)
'
'        ElseIf SELDM_RS!PRO_CODE = "F894" Then
'            UPDBIO_QR(34) = Trim(SELDM_RS!OBSERVALUE)
'
'        ElseIf SELDM_RS!PRO_CODE = "F893" Then
'            UPDBIO_QR(35) = Trim(SELDM_RS!OBSERVALUE)
'
'        ElseIf SELDM_RS!PRO_CODE = "F845" Then
'            UPDBIO_QR(36) = Trim(SELDM_RS!OBSERVALUE)
'        End If
'
'        SELDM_RS.MoveNext
'     Loop
'
'    UPDBIO_QR(37) = Trim(i)
'    UPDBIO_QR.Execute
'    selreq_rs.MoveNext
'Loop
'CR.Formulas(0) = "CASENO='" & Msk_caseno.Text & "'"
''CR.ReportFileName = "C:\DISREPORTSEC\BIOCHEMICAL2.rpt"
'CR.ReportFileName = "C:\DISREPORTSEC\BIOCHEMICALFinalreport.rpt"
'CR.Action = 1
'End If
'
End If

''for Tumor Marker
If mod1 = "TM" Then


     If OPT_1.Value = True Then
        INSDRUGREG_QR(0) = GET_REVDATE(msk_fromdt.Text)
        INSDRUGREG_QR(1) = GET_REVDATE(msk_todt.Text)
        INSDRUGREG_QR.Execute
     
        
        'Set SELDRUGREG_RS = INSDRUGREG_QR.OpenResultset(1, 3)
        Set SELDRUGREG_RS = SELDRUGREG_QR.OpenResultset(1, 3)
        
        Do Until SELDRUGREG_RS.EOF
            DoEvents
            i = SELDRUGREG_RS!REQNO
        
            seldmulproce_qr(0) = Trim(i)
            seldmulproce_qr.Execute
            Set seldmulproce_rs = seldmulproce_qr.OpenResultset(1, 3)
        
            UPDDRUGREG_QR(0) = ""
            UPDDRUGREG_QR(1) = ""
            UPDDRUGREG_QR(2) = ""
            UPDDRUGREG_QR(3) = ""
            UPDDRUGREG_QR(4) = ""
        
            Do Until seldmulproce_rs.EOF
                If seldmulproce_rs!PRO_CODE = "F834" Then
                    UPDDRUGREG_QR(0) = "YES"
                ElseIf seldmulproce_rs!PRO_CODE = "F833" Then
                    UPDDRUGREG_QR(1) = "YES"
                ElseIf seldmulproce_rs!PRO_CODE = "F835" Then
                    UPDDRUGREG_QR(2) = "YES"
                ElseIf seldmulproce_rs!PRO_CODE = "F836" Then
                    UPDDRUGREG_QR(3) = "YES"
                End If
                  
                seldmulproce_rs.MoveNext
            Loop
        
            SELDOSE_QR(0) = Trim(i)
            Set SELDOSE_RS = SELDOSE_QR.OpenResultset(1, 3)
            
            UPDDRUGREG_QR(4) = SELDOSE_RS!PREPATHNO
            UPDDRUGREG_QR(5) = Trim(i)
            UPDDRUGREG_QR.Execute
        
            SELDRUGREG_RS.MoveNext
        
        Loop
        
        CR.Formulas(0) = "FRMDT= '" & msk_fromdt.Text & "'"
        CR.Formulas(1) = "TODT='" & msk_todt.Text & "'"
        CR.ReportFileName = "C:\DISREPORTSEC\DRUGREG.rpt"
        CR.WindowState = crptMaximized
        CR.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CR.Action = 1
        
        DELDRUGREG_QR.Execute
        
     End If
    
    If OPT_2.Value = True Then
        SELTM_QR(0) = GET_REVDATE(msk_fromdt.Text)
        SELTM_QR(1) = GET_REVDATE(msk_todt.Text)
        SELTM_QR.Execute
    
        Set SELTM_RS = seltumtemp_qr.OpenResultset(1, 3)
    
        Do Until SELTM_RS.EOF
        DoEvents
        i = SELTM_RS!REQNO
    
        SELDM_QR(0) = i
        SELDM_QR.Execute
        Set SELDM_RS = SELDM_QR.OpenResultset(1, 3)
    
        updtumtemp_qr(0) = ""
        updtumtemp_qr(1) = ""
        updtumtemp_qr(2) = ""
        updtumtemp_qr(3) = ""
        updtumtemp_qr(4) = ""
        updtumtemp_qr(5) = ""
        updtumtemp_qr(6) = ""
        updtumtemp_qr(7) = ""
        updtumtemp_qr(8) = ""
        updtumtemp_qr(9) = ""
        updtumtemp_qr(10) = ""
        updtumtemp_qr(11) = ""
        updtumtemp_qr(12) = ""
        updtumtemp_qr(13) = ""
        updtumtemp_qr(14) = ""
        updtumtemp_qr(15) = ""
        updtumtemp_qr(16) = ""
        updtumtemp_qr(17) = ""
        updtumtemp_qr(18) = ""
        updtumtemp_qr(19) = ""
        updtumtemp_qr(20) = ""
        updtumtemp_qr(21) = ""
        updtumtemp_qr(22) = ""
        updtumtemp_qr(23) = ""
        updtumtemp_qr(24) = ""
        updtumtemp_qr(25) = ""
        updtumtemp_qr(26) = ""
        updtumtemp_qr(27) = ""
        updtumtemp_qr(28) = ""
'            Do Until SELDM_RS.EOF
'        If SELDM_RS!PRO_CODE = "F817" Then
'            updtumtemp_qr(0) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F818" Then
'            updtumtemp_qr(1) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F819" Then
'            updtumtemp_qr(2) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(3) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(4) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(5) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(6) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(7) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(8) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(9) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(10) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(11) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(12) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(13) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(14) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(15) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(16) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(17) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(18) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F820" Then
'            updtumtemp_qr(19) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F977" Then
'            updtumtemp_qr(20) = Trim(SELDM_RS!OBSERVALUE)
'        End If
'
'        SELDM_RS.MoveNext
'     Loop
    '***** Commented and added by Shraddha on 11/05/10 *****
        Do Until SELDM_RS.EOF
        If SELDM_RS!PRO_CODE = "F817" Then
            updtumtemp_qr(0) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F818" Then
            updtumtemp_qr(1) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F819" Then
            updtumtemp_qr(2) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F820" Then
            updtumtemp_qr(3) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F821" Then
            updtumtemp_qr(4) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F822" Then
            updtumtemp_qr(5) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F823" Then
            updtumtemp_qr(6) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F824" Then
            updtumtemp_qr(7) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F826" Then
            updtumtemp_qr(8) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F827" Then
            updtumtemp_qr(9) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F828" Then
            updtumtemp_qr(10) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F829" Then
            updtumtemp_qr(11) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F830" Then
            updtumtemp_qr(12) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F831" Then
            updtumtemp_qr(13) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F832" Then
            updtumtemp_qr(14) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F833" Then
            updtumtemp_qr(15) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F834" Then
            updtumtemp_qr(16) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F835" Then
            updtumtemp_qr(17) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F836" Then
            updtumtemp_qr(18) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F843" Then
            updtumtemp_qr(19) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F977" Then
            updtumtemp_qr(20) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F038" Then
            updtumtemp_qr(21) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F034" Then
            updtumtemp_qr(22) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F035" Then
            updtumtemp_qr(23) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F036" Then
            updtumtemp_qr(24) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F037" Then
            updtumtemp_qr(25) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F060" Then
            updtumtemp_qr(26) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F058" Then
            updtumtemp_qr(27) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F059" Then
            updtumtemp_qr(28) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F837" Then
            updtumtemp_qr(29) = Trim(SELDM_RS!OBSERVALUE)
        
        End If
        
        
        SELDM_RS.MoveNext
     Loop
    
        updtumtemp_qr(30) = Trim(i)
        updtumtemp_qr.Execute
        SELTM_RS.MoveNext
        
    Loop

        CR.Formulas(0) = "FRMDT= '" & msk_fromdt.Text & "'"
        CR.Formulas(1) = "TODT='" & msk_todt.Text & "'"
        'CR.ReportFileName = "C:\DISREPORTSEC\BIOCHEMICAL2.rpt"
        CR.ReportFileName = "C:\DISREPORTSEC\TUMORMARKER.rpt"
        CR.WindowState = crptMaximized
        CR.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CR.Action = 1

    End If
    

    If OPT_3.Value = True Then
    
        CR1.ReportFileName = "C:\DISREPORTSEC\ELECTROREG.rpt"
        CR1.Formulas(0) = "FRMDT= '" & msk_fromdt.Text & "'"
        CR1.Formulas(1) = "TODT='" & msk_todt.Text & "'"
        CR1.SelectionFormula = "{RREQUISITN.PREPORTDT}>=Date(" & var_date_f & ") AND {RREQUISITN.PREPORTDT}<=Date(" & var_date_t & ")"
        CR1.WindowState = crptMaximized
        CR1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CR1.Action = 1
        
    
    End If


End If

'SHWETA 06-JULY-2010

'' for Emergency Lab Report

If mod1 = "EL" Then

'If Option1.Value = True Then
    SELEL_QR(0) = GET_REVDATE(msk_fromdt.Text)
    SELEL_QR(1) = GET_REVDATE(msk_todt.Text)
    SELEL_QR.Execute
    
    Set SELELREQ_RS = SELELREQ_QR.OpenResultset(1, 3)
    Do Until SELELREQ_RS.EOF
    DoEvents
    i = Trim(SELELREQ_RS!REQNO)
    SELDM_QR(0) = i
    SELDM_QR.Execute
    Set SELDM_RS = SELDM_QR.OpenResultset(1, 3)
    UPDEL_QR(0) = ""
    UPDEL_QR(1) = ""
    UPDEL_QR(2) = ""
    UPDEL_QR(3) = ""
    UPDEL_QR(4) = ""
    UPDEL_QR(5) = ""
    UPDEL_QR(6) = ""
    UPDEL_QR(7) = ""
    UPDEL_QR(8) = ""
    UPDEL_QR(9) = ""
    UPDEL_QR(10) = ""
    UPDEL_QR(11) = ""
    UPDEL_QR(12) = ""
    UPDEL_QR(13) = ""
    UPDEL_QR(14) = ""
    UPDEL_QR(15) = ""
    UPDEL_QR(16) = ""
    UPDEL_QR(17) = ""
    UPDEL_QR(18) = ""
    UPDEL_QR(19) = ""
    UPDEL_QR(20) = ""
    UPDEL_QR(21) = ""
    UPDEL_QR(22) = ""
    UPDEL_QR(23) = ""
    UPDEL_QR(24) = ""
    UPDEL_QR(25) = ""
    UPDEL_QR(26) = ""
    UPDEL_QR(27) = ""
    UPDEL_QR(28) = ""
    UPDEL_QR(29) = ""
'    UPDEL_QR(30) = ""
'    UPDEL_QR(31) = ""
'    UPDEL_QR(32) = ""
'    UPDEL_QR(33) = ""
'    UPDEL_QR(34) = ""
'    UPDEL_QR(35) = ""
'    UPDEL_QR(36) = ""
    
    Do Until SELDM_RS.EOF
    ' F031= 0?,F032=1?,F848=2?,F877=3?,F878=4?,
        If SELDM_RS!PRO_CODE = "F031" Then
            UPDEL_QR(0) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F032" Then
            UPDEL_QR(1) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F848" Then
            UPDEL_QR(2) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F877" Then
            UPDEL_QR(3) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F878" Then
            UPDEL_QR(4) = Trim(SELDM_RS!OBSERVALUE)
            'F879=5?, F911=6?,F912=7?,F914=8?,F950=9?,
        ElseIf SELDM_RS!PRO_CODE = "F879" Then
            UPDEL_QR(5) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F911" Then
            UPDEL_QR(6) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F912" Then
            UPDEL_QR(7) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F914" Then
            UPDEL_QR(8) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F950" Then
            UPDEL_QR(9) = Trim(SELDM_RS!OBSERVALUE)
            'F955=10?,F956=11?,F957=12?,F958=13?,F961=14?,F978=15?,
        ElseIf SELDM_RS!PRO_CODE = "F955" Then
            UPDEL_QR(10) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F956" Then
            UPDEL_QR(11) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F957" Then
            UPDEL_QR(12) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F958" Then
            UPDEL_QR(13) = Trim(SELDM_RS!OBSERVALUE)
             
        ElseIf SELDM_RS!PRO_CODE = "F961" Then
            UPDEL_QR(14) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F978" Then
            UPDEL_QR(15) = Trim(SELDM_RS!OBSERVALUE)
            'F979=16?,F980=17?,F982=18?,F983=19?,F984=20?,F985=21?
        ElseIf SELDM_RS!PRO_CODE = "F979" Then
            UPDEL_QR(16) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F980" Then
            UPDEL_QR(17) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F982" Then
            UPDEL_QR(18) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F983" Then
            UPDEL_QR(19) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F984" Then
            UPDEL_QR(20) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F985" Then
            UPDEL_QR(21) = Trim(SELDM_RS!OBSERVALUE)
            'F986=22?,F987=23?,F988=24?,F990=25?,F991=26?,
        ElseIf SELDM_RS!PRO_CODE = "F986" Then
            UPDEL_QR(22) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F987" Then
            UPDEL_QR(23) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F988" Then
            UPDEL_QR(24) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F990" Then
            UPDEL_QR(25) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F991" Then
            UPDEL_QR(26) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F993" Then
            UPDEL_QR(27) = Trim(SELDM_RS!OBSERVALUE)
        ElseIf SELDM_RS!PRO_CODE = "F998" Then
            UPDEL_QR(28) = Trim(SELDM_RS!OBSERVALUE)
       
        End If
        
        SELDM_RS.MoveNext
     Loop
    
    UPDEL_QR(29) = Trim(i)
    UPDEL_QR.Execute
    SELELREQ_RS.MoveNext
Loop
CR.Formulas(0) = "FRMDT= '" & msk_fromdt.Text & "'"
CR.Formulas(1) = "TODT='" & msk_todt.Text & "'"
'CR.ReportFileName = "C:\DISREPORTSEC\BIOCHEMICAL2.rpt"
CR.ReportFileName = "C:\DISREPORTSEC\EmergencyLabFinalreport.rpt"
CR.Action = 1
'End If

'If Option2.Value = True Then
'    SELRREQ_QR(0) = Trim(Msk_caseno.Text)
'    SELRREQ_QR.Execute
'    Set selreq_rs = selreq_qr.OpenResultset(1, 3)
'    Do Until selreq_rs.EOF
'    i = selreq_rs!reqno
'    SELDM_QR(0) = i
'    SELDM_QR.Execute
'    Set SELDM_RS = SELDM_QR.OpenResultset(1, 3)
'    UPDBIO_QR(0) = ""
'    UPDBIO_QR(1) = ""
'    UPDBIO_QR(2) = ""
'    UPDBIO_QR(3) = ""
'    UPDBIO_QR(4) = ""
'    UPDBIO_QR(5) = ""
'    UPDBIO_QR(6) = ""
'    UPDBIO_QR(7) = ""
'    UPDBIO_QR(8) = ""
'    UPDBIO_QR(9) = ""
'    UPDBIO_QR(10) = ""
'    UPDBIO_QR(11) = ""
'    UPDBIO_QR(12) = ""
'    UPDBIO_QR(13) = ""
'    UPDBIO_QR(14) = ""
'    UPDBIO_QR(15) = ""
'    UPDBIO_QR(16) = ""
'    UPDBIO_QR(17) = ""
'    UPDBIO_QR(18) = ""
'    UPDBIO_QR(19) = ""
'    UPDBIO_QR(20) = ""
'    UPDBIO_QR(21) = ""
'    UPDBIO_QR(22) = ""
'    UPDBIO_QR(23) = ""
'    UPDBIO_QR(24) = ""
'    UPDBIO_QR(25) = ""
'    UPDBIO_QR(26) = ""
'    UPDBIO_QR(27) = ""
'    UPDBIO_QR(28) = ""
'    UPDBIO_QR(29) = ""
'    UPDBIO_QR(30) = ""
'    UPDBIO_QR(31) = ""
'    UPDBIO_QR(32) = ""
'    UPDBIO_QR(33) = ""
'    UPDBIO_QR(34) = ""
'    UPDBIO_QR(35) = ""
'    UPDBIO_QR(36) = ""
'
'    Do Until SELDM_RS.EOF
'        If SELDM_RS!PRO_CODE = "F842" Then
'            UPDBIO_QR(0) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F843" Then
'            UPDBIO_QR(1) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F841" Then
'            UPDBIO_QR(2) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F846" Then
'            UPDBIO_QR(3) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F847" Then
'            UPDBIO_QR(4) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F855" Then
'            UPDBIO_QR(5) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F856" Then
'            UPDBIO_QR(6) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F857" Then
'            UPDBIO_QR(7) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F861" Then
'            UPDBIO_QR(8) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F862" Then
'            UPDBIO_QR(9) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F863" Then
'            UPDBIO_QR(10) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F864" Then
'            UPDBIO_QR(11) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F888" Then
'            UPDBIO_QR(12) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F890" Then
'            UPDBIO_QR(13) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F866" Then
'            UPDBIO_QR(14) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F867" Then
'            UPDBIO_QR(15) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F868" Then
'            UPDBIO_QR(16) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F869" Then
'            UPDBIO_QR(17) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F870" Then
'            UPDBIO_QR(18) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F871" Then
'            UPDBIO_QR(19) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F872" Then
'            UPDBIO_QR(20) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F873" Then
'            UPDBIO_QR(21) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F874" Then
'            UPDBIO_QR(22) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F876" Then
'            UPDBIO_QR(23) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F895" Then
'            UPDBIO_QR(24) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F896" Then
'            UPDBIO_QR(25) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F850" Then
'            UPDBIO_QR(26) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F853" Then
'            UPDBIO_QR(27) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F851" Then
'            UPDBIO_QR(28) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F852" Then
'            UPDBIO_QR(29) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F881" Then
'            UPDBIO_QR(30) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F882" Then
'            UPDBIO_QR(31) = Trim(SELDM_RS!OBSERVALUE)
'        ElseIf SELDM_RS!PRO_CODE = "F891" Then
'            UPDBIO_QR(32) = Trim(SELDM_RS!OBSERVALUE)
'
'        ElseIf SELDM_RS!PRO_CODE = "F924" Then
'            UPDBIO_QR(33) = Trim(SELDM_RS!OBSERVALUE)
'
'        ElseIf SELDM_RS!PRO_CODE = "F894" Then
'            UPDBIO_QR(34) = Trim(SELDM_RS!OBSERVALUE)
'
'        ElseIf SELDM_RS!PRO_CODE = "F893" Then
'            UPDBIO_QR(35) = Trim(SELDM_RS!OBSERVALUE)
'
'        ElseIf SELDM_RS!PRO_CODE = "F845" Then
'            UPDBIO_QR(36) = Trim(SELDM_RS!OBSERVALUE)
'        End If
'
'        SELDM_RS.MoveNext
'     Loop
'
'    UPDBIO_QR(37) = Trim(i)
'    UPDBIO_QR.Execute
'    selreq_rs.MoveNext
'Loop
'CR.Formulas(0) = "CASENO='" & Msk_caseno.Text & "'"
''CR.ReportFileName = "C:\DISREPORTSEC\BIOCHEMICAL2.rpt"
'CR.ReportFileName = "C:\DISREPORTSEC\BIOCHEMICALFinalreport.rpt"
'CR.Action = 1
'End If
'
End If






''''''TILL HERE

 TXT_MESSAGE.Text = "  Click on BACK button to EXIT......"
    MYDELAY (5)


CMDREPORT.Enabled = False
cmd_cancel.Enabled = True
Me.MousePointer = 0

End Sub

Private Sub Form_Activate()
TXT_MESSAGE.Text = ""
End Sub

Private Sub Form_Load()
Set RS = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107") '***2/2/99
DATE11 = RS("dt1")

str = "INSERT INTO DISLIB.BIOTEMP(REQNO,DATE1,CASENO,LABNO,SAMPLE,NAME)" _
& " SELECT A.REQNNO,A.REQNDT,A.CASENO,B.PATHNO,B.NMETERIALS,C.FIRST_NAME " _
& " FROM RISLIB.RREQUISITN A,DISLIB.DREQUISITN B,PATOOLSLIB.TA101 C " _
& " WHERE A.REQNNO=B.REQNO AND A.CASENO=C.CASE_NO  AND A.CASENO =? "

Set SELRREQ_QR = rdb.CreateQuery("RR", str)

str = "INSERT INTO DISLIB.BIOTEMP(REQNO,DATE1,CASENO,LABNO,SAMPLE,NAME)" _
& " SELECT A.REQNNO,A.REQNDT,A.CASENO,B.PATHNO,B.NMETERIALS,C.FIRST_NAME " _
& " FROM RISLIB.RREQUISITN A,DISLIB.DREQUISITN B,PATOOLSLIB.TA101 C " _
& " WHERE A.REQNNO=B.REQNO AND A.CASENO=C.CASE_NO  AND A.CASENO =? "

Set SELTM_QR1 = rdb.CreateQuery("TT", str)


str = "INSERT INTO DISLIB.BIOTEMP(REQNO,DATE1,CASENO,LABNO,SAMPLE,NAME)" _
& " SELECT A.REQNNO,A.REQNDT,A.CASENO,B.PATHNO,B.NMETERIALS,C.FIRST_NAME " _
& " FROM RISLIB.RREQUISITN A,DISLIB.DREQUISITN B,PATOOLSLIB.TA101 C " _
& " WHERE A.REQNNO=B.REQNO AND A.CASENO=C.CASE_NO AND A.PREPORTDT BETWEEN ? AND ?  AND A.MODALITYCD='BC' AND (A.REQNSTAT='T' OR A.REQNSTAT='F')"
Set SELRREQ_QR1 = rdb.CreateQuery("RR1", str)


str = "INSERT INTO DISLIB.ELTEMP(REQNO,DATE1,CASENO,LABNO,SAMPLE,NAME)" _
& " SELECT A.REQNNO,A.REQNDT,A.CASENO,B.PATHNO,B.NMETERIALS,C.FIRST_NAME " _
& " FROM RISLIB.RREQUISITN A,DISLIB.DREQUISITN B,PATOOLSLIB.TA101 C " _
& " WHERE A.REQNNO=B.REQNO AND A.CASENO=C.CASE_NO AND A.PREPORTDT BETWEEN ? AND ?  AND A.MODALITYCD='EL' AND (A.REQNSTAT='T' OR A.REQNSTAT='F')"
Set SELEL_QR = rdb.CreateQuery("EL1", str)

str = "INSERT INTO DISLIB.TUMTEMP(REQNO,DATE1,CASENO,LABNO,SAMPLE,NAME)" _
& " SELECT A.REQNNO,A.REQNDT,A.CASENO,B.PATHNO,B.NMETERIALS,C.FIRST_NAME " _
& " FROM RISLIB.RREQUISITN A,DISLIB.DREQUISITN B,PATOOLSLIB.TA101 C " _
& " WHERE A.REQNNO=B.REQNO AND A.CASENO=C.CASE_NO AND A.PREPORTDT BETWEEN ? AND ?  AND A.MODALITYCD='TM' AND (A.REQNSTAT='T' OR A.REQNSTAT='F')"

Set SELTM_QR = rdb.CreateQuery("TM1", str)

'STR = " SELECT A.REQNNO,A.REQNDT,A.CASENO,B.PATHNO,B.NMETERIALS,C.FIRST_NAME " _
'& " FROM RISLIB.RREQUISITN A,DISLIB.DREQUISITN B,PATOOLSLIB.TA101 C " _
'& " WHERE A.REQNNO=B.REQNO AND A.CASENO=C.CASE_NO  AND A.PREPORTDT BETWEEN ? AND ?  AND A.MODALITYCD='BC' AND (A.REQNSTAT='T' OR A.REQNSTAT='F')"
'Set SELRREQ_QR1 = rdb.CreateQuery("RR1", STR)

'THIS IS FOR INSERT RECORDS IN BIOTEMP
'STR = "INSERT INTO DISLIB.BIOTEMP(REQNO,CASENO,LABNO,SAMPLE,NAME) VALUES (?,?,?,?,?) "
'Set INSBIOTEMP_QR = rdb.CreateQuery("INSBIO", STR)

str = "DELETE FROM DISLIB.BIOTEMP"
Set DEL_QR = rdb.CreateQuery("DD", str)

str = "DELETE FROM DISLIB.TUMTEMP"
Set DELTUMTEMP_QR = rdb.CreateQuery("DELETETUMPTEMP", str)

str = "DELETE FROM DISLIB.ELTEMP"
Set DELELTEMP_QR = rdb.CreateQuery("DELETEELTEMP", str)

str = "SELECT PRO_CODE,OBSERVALUE FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
Set SELDM_QR = rdb.CreateQuery("DM", str)

str = "SELECT REQNO FROM DISLIB.BIOTEMP"
Set selreq_qr = rdb.CreateQuery("REQ", str)
'SHWETA 6-JULY-2010
str = "SELECT REQNO FROM ELTEMP"
Set SELELREQ_QR = rdb.CreateQuery("ELREQ", str)
'''''''''
str = "SELECT * FROM DISLIB.TUMTEMP"
Set seltumtemp_qr = rdb.CreateQuery("TUMTEMP", str)

str = "UPDATE DISLIB.BIOTEMP SET F842= ?,F843=?,F841=?,F846=?,F847=?," _
& " F855=?, F856=?,F857=?,F861=?,F862=?,F863=?,F864=?,F888=?,F890=?,F866=?,F867=?, " _
& " F868=?,F869=?,F870=?,F871=?,F872=?,F873=?,F874=?,F876=?,F895=?,F896=?,F850=?, " _
& " F853=?,F851=?,F852=?,F881=?,F882=?,F891=?,F924=?,F894=?,F893=?,F845=? WHERE REQNO =?"
Set UPDBIO_QR = rdb.CreateQuery("UPD", str)

'SHWETA 6-JULY-2010 FOR EL
str = "UPDATE DISLIB.ELTEMP SET F031= ?,F032=?,F848=?,F877=?,F878=?," _
& " F879=?, F911=?,F912=?,F914=?,F950=?,F955=?,F956=?,F957=?,F958=?,F961=?,F978=?, " _
& " F979=?,F980=?,F982=?,F983=?,F984=?,F985=?,F986=?,F987=?,F988=?,F990=?,F991=?, " _
& " F993=?,F998=? WHERE REQNO =?"
Set UPDEL_QR = rdb.CreateQuery("UPDEL", str)


'str = "UPDATE DISLIB.TUMTEMP SET F817= ?,F818= ?,F819= ?,F820= ?,F821= ?,F822= ?,F823= ?,F824= ?," _
'& "F826= ?,F827= ?,F828= ?,F829= ?,F830= ?,F831= ?,F832= ?,F833= ?,F834= ?,F835= ?,F836= ?,F843= ?,F977= ? WHERE REQNO =?"
'Set updtumtemp_qr = rdb.CreateQuery("UPDATE1", str)

'***** Commented and added by Shraddha 11/05/10 *****
str = "UPDATE DISLIB.TUMTEMP SET F817= ?,F818= ?,F819= ?,F820= ?,F821= ?,F822= ?,F823= ?,F824= ?," _
& "F826= ?,F827= ?,F828= ?,F829= ?,F830= ?,F831= ?,F832= ?,F833= ?,F834= ?,F835= ?,F836= ?,F843= ?,F977= ?,FLD1=?, FLD2=?, FLD3=?, FLD4=?, F037=?, F060=?, F058=?, F059=?, F837=?, WHERE REQNO =?"
Set updtumtemp_qr = rdb.CreateQuery("UPDATE1", str)
'FLD1 FLD2 FLD3 FLD4
'****************************************************

str = "DELETE FROM DISLIB.TUMTEMP"
Set DELTUMTEMP_QR = rdb.CreateQuery("DELETE1", str)

'''DRUGREGISTER

'STR = "INSERT INTO DISLIB.DRUGREGTMP(REQNO,LABNO)" _
'& " SELECT A.REQNNO,B.PATHNO" _
'& " FROM RISLIB.RREQUISITN A,DISLIB.DREQUISITN B " _
'& " WHERE A.REQNNO=B.REQNO AND A.PREPORTDT BETWEEN ? AND ?  AND A.MODALITYCD='TM' AND (A.REQNSTAT='T' OR A.REQNSTAT='F')"

str = "INSERT INTO DISLIB.DRUGREGTMP(REQNO,LABNO)" _
& " SELECT A.REQNNO,B.PATHNO" _
& " FROM RISLIB.RREQUISITN A,DISLIB.DREQUISITN B,DISLIB.DMULPROCE C " _
& " WHERE A.REQNNO=B.REQNO AND A.REQNNO=C.REQNO AND A.PREPORTDT BETWEEN ? AND ?  AND A.MODALITYCD='TM' AND (A.REQNSTAT='T' OR A.REQNSTAT='F') AND C.PRO_CODE IN('F834','F833','F835','F836')"
Set INSDRUGREG_QR = rdb.CreateQuery("INSDRUGREGTMP", str)


str = "SELECT * FROM DISLIB.DRUGREGTMP"
Set SELDRUGREG_QR = rdb.CreateQuery("SELECTDRUGREG", str)

str = "UPDATE DISLIB.DRUGREGTMP SET AMIKACIN=?,CYCLOIN=?,VACO=?,METHO=?,DOSE=? WHERE REQNO=?"
Set UPDDRUGREG_QR = rdb.CreateQuery("UPDATEDREUGREG", str)

str = "SELECT PRO_CODE FROM DISLIB.DMULPROCE WHERE REQNO=?"
Set seldmulproce_qr = rdb.CreateQuery("SELDMUL", str)

str = "SELECT PREPATHNO FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELDOSE_QR = rdb.CreateQuery("SELECTDOSE", str)

str = "DELETE FROM DISLIB.DRUGREGTMP"
Set DELDRUGREG_QR = rdb.CreateQuery("DELDRUGREGTMP", str)



cmb_mod.AddItem "Biochemistry$BC"
cmb_mod.AddItem "Tumor Marker$TM"
cmb_mod.AddItem "Emergency Lab$EL"

'CMDREPORT.Enabled = False
Frame4.Enabled = False
TXT_MESSAGE = ""

End Sub

'Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
' DEL_QR.Execute
' DELTUMTEMP_QR.Execute
'End Sub

Private Sub msk_caseno_LostFocus()
msk_caseno.Text = UCase(msk_caseno.Text)
'If Option2.Value = True Then
'If Trim(Msk_caseno.Text) = "" Then
'   MsgBox "Case No. is Empty"
'Exit Sub
'End If
'  CASENO = padzero(5, "C", Msk_caseno.Text, 3)
'  Msk_caseno.Text = UCase(CASENO)
'End If
End Sub

Private Sub Msk_fromdt_GotFocus()
    msk_fromdt.Text = Format(DATE11, "DD/MM/YYYY")
End Sub

Private Sub Msk_todt_GotFocus()
    msk_todt.Text = Format(DATE11, "DD/MM/YYYY")
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




