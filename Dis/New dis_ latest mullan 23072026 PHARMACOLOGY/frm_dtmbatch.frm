VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_dtmbatch 
   Caption         =   "Form2"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   LinkTopic       =   "Form2"
   ScaleHeight     =   5700
   ScaleWidth      =   12615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   1560
      TabIndex        =   11
      Top             =   7200
      Visible         =   0   'False
      Width           =   9135
      Begin VB.ComboBox cmb_ward 
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
         Height          =   315
         Left            =   6000
         TabIndex        =   18
         Top             =   1680
         Width           =   2775
      End
      Begin VB.ComboBox cmb_patstat 
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
         Left            =   1920
         TabIndex        =   17
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox cmb_patcatg 
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
         Left            =   6000
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cmb_pattype 
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
         Left            =   1920
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "Back"
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
         Left            =   5760
         TabIndex        =   14
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "Print"
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
         Left            =   2280
         TabIndex        =   13
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmb_ref 
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
         Left            =   3840
         TabIndex        =   12
         Top             =   2280
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Msk_todt 
         Height          =   375
         Left            =   6000
         TabIndex        =   19
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Msk_fromdt 
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   0
         PromptInclude   =   0   'False
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
      Begin VB.Label Label7 
         Caption         =   "Ward"
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
         Left            =   5400
         TabIndex        =   26
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Patient Status"
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
         Left            =   600
         TabIndex        =   25
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Patient Category"
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
         Left            =   4440
         TabIndex        =   24
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Patient Type"
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
         Left            =   720
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label LBL_TODT 
         AutoSize        =   -1  'True
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
         Left            =   5160
         TabIndex        =   22
         Top             =   360
         Width           =   705
      End
      Begin VB.Label LBL_FROMDT 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   21
         Top             =   360
         Width           =   885
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   2040
      TabIndex        =   0
      Top             =   1440
      Width           =   6855
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton Command1 
         Caption         =   "E&xit"
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
         Left            =   5040
         TabIndex        =   8
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
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
         Left            =   3480
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton CMD_PRINT1 
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
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmd_ref 
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
         Left            =   1920
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin MSMask.MaskEdBox MSK_TODT1 
         Height          =   375
         Left            =   4920
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox MSK_FROMDT1 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   0
         PromptInclude   =   0   'False
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
      Begin VB.Label Label9 
         Caption         =   "To Date "
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
         Left            =   3960
         TabIndex        =   33
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "From Time "
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
         Left            =   480
         TabIndex        =   32
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label LBL_FROMDT1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   885
      End
      Begin VB.Label LBL_TODT1 
         AutoSize        =   -1  'True
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
         Left            =   3960
         TabIndex        =   9
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Label Label3 
      Caption         =   "05/11/21"
      Height          =   495
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label LBL_CAP 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "DTM BATCH PRINTING"
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
      Left            =   3840
      TabIndex        =   30
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label1 
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
      Left            =   600
      TabIndex        =   29
      Top             =   3120
      Width           =   5055
   End
   Begin VB.Label LBL_PRINTED 
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   960
      TabIndex        =   28
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label2 
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
      Left            =   840
      TabIndex        =   27
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "frm_dtmbatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************
'This Form has entirely been developed by Ms. Shraddha Pandit on 03/03/09
'************************************************************************

Dim str As String

Dim selpreqnoqr As rdoQuery
Dim selpreqnors As rdoResultset

Dim SELDTMNO_RS As rdoResultset
Dim SELDTMNO_QR As rdoQuery
Dim SELDTMTAB_QR As rdoQuery
Dim SELDTMTAB_RS As rdoResultset

Dim SELEMPBASINFO_QR As rdoQuery
Dim SELEMPBASINFO_RS As rdoResultset
Dim req2_qr As rdoQuery
Dim req2_rs As rdoResultset
Dim DrequisitnQR As rdoQuery
Dim DrequisitnRS As rdoResultset
Dim UPD_DTMTAB_QR As rdoQuery

Dim frm_tm, to_tm As String
Dim hr, mn
Dim sel_repttm_qr As rdoQuery
Dim sel_repttm_rs As rdoResultset


Private Sub CMD_PRINT1_Click()
Dim CC As Integer
Dim a As Integer
CC = 0
Dim p_count As Integer
On Error GoTo MSG:

selpreqnoqr.QueryTimeout = 1000

'********** Shraddha 01/09/09 **********
hr = Left(Trim(MaskEdBox1.FormattedText), 2)
mn = Right(Trim(MaskEdBox1.FormattedText), 2)
If hr > 24 Then
    MsgBox "Please enter valid FROM time"
    MaskEdBox1.SetFocus
    MaskEdBox1.Text = ""
    Exit Sub
ElseIf mn > 59 Then
    MsgBox "Please enter valid FROM time"
    MaskEdBox1.SetFocus
    MaskEdBox1.Text = ""
    Exit Sub
End If

hr = Left(Trim(MaskEdBox2.FormattedText), 2)
mn = Right(Trim(MaskEdBox2.FormattedText), 2)
If hr > 24 Then
    MsgBox "Please enter valid TO time"
    MaskEdBox2.SetFocus
    MaskEdBox2.Text = ""
    Exit Sub
ElseIf mn > 59 Then
    MsgBox "Please enter valid TO time"
    MaskEdBox2.SetFocus
    MaskEdBox2.Text = ""
    Exit Sub
End If
'**************************************

frm_tm = Trim(MaskEdBox1.FormattedText) & ":00"
to_tm = Trim(MaskEdBox2.FormattedText) & ":00"

'*********
'selpreqnoqr(0) = Trim(MSK_FROMDT1.FormattedText)
'selpreqnoqr(1) = Trim(MSK_TODT1.FormattedText)
'above 2 original 31/08/09


selpreqnoqr(0) = frm_tm
selpreqnoqr(1) = to_tm
selpreqnoqr(2) = Trim(MSK_FROMDT1.FormattedText)
selpreqnoqr(3) = Trim(MSK_TODT1.FormattedText)
'*********
Set selpreqnors = selpreqnoqr.OpenResultset(1, 3)
p_count = 0
'If not selpreqnors.RowCount > 0 Then
If Not selpreqnors.EOF Then
         
    Label1.Caption = "No. of Reports in this Combination are " & selpreqnors.RowCount
    LBL_PRINTED.Caption = CC
   
    While Not selpreqnors.EOF
        CC = CC + 1
        
        If CC = 1 Or CC = 21 Or CC = 31 Or CC = 41 Or CC = 51 Or CC = 61 Or CC = 71 Or CC = 81 Or CC = 91 Then
            LBL_PRINTED.Caption = "Now processing " & CC & "st report"
        ElseIf CC = 2 Or CC = 22 Or CC = 32 Or CC = 42 Or CC = 52 Or CC = 62 Or CC = 72 Or CC = 82 Or CC = 92 Then
            LBL_PRINTED.Caption = "Now processing " & CC & "nd report"
        ElseIf CC = 3 Or CC = 23 Or CC = 33 Or CC = 43 Or CC = 53 Or CC = 63 Or CC = 73 Or CC = 83 Or CC = 93 Then
            LBL_PRINTED.Caption = "Now processing " & CC & "rd report"
        Else
            LBL_PRINTED.Caption = "Now processing " & CC & "th report"
        End If
        
        DrequisitnQR(0) = selpreqnors("REQNnO")
        Set DrequisitnRS = DrequisitnQR.OpenResultset(1, 3)
    
        SELDTMNO_QR(0) = DrequisitnRS("CASENO")
        Set SELDTMNO_RS = SELDTMNO_QR.OpenResultset(1, 3)
        
        SELDTMTAB_QR(0) = selpreqnors("REQNnO")
        Set SELDTMTAB_RS = SELDTMTAB_QR.OpenResultset(1, 3)
        
        'Check whether it print was taken earlier
        If SELDTMTAB_RS("HBSAGKEY") <> 1 Then
                
            If Len(selpreqnors("CASENO")) = 6 Then 'Staff
            
                SELEMPBASINFO_QR(0) = selpreqnors("CASENO")
                Set SELEMPBASINFO_RS = SELEMPBASINFO_QR.OpenResultset(1, 3)
                   
                req2_qr(0) = selpreqnors("REQNnO")
                Set req2_rs = req2_qr.OpenResultset(1, 3)
                
                str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & selpreqnors("REQNnO") & "'"
                'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & selpreqnors("REQNnO") & "' and {DTMVIEW1.REPTTM} >='" & frm_tm & "' and {DTMVIEW1.REPTTM} <='" & to_tm & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtmstaff.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & selpreqnors("REQNnO") & "'"
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
                End If
                
                CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(3) = "NAM= '" & SELEMPBASINFO_RS("Title") & " " & SELEMPBASINFO_RS("FNAME") & " " & SELEMPBASINFO_RS("MNAME") & " " & SELEMPBASINFO_RS("LNAME") & "'"
                CrystalReport1.Formulas(4) = "DOB= '" & DateDiff("YYYY", SELEMPBASINFO_RS("dob"), req2_rs("reqndt")) & "'"
                CrystalReport1.Formulas(5) = "STDOB= '" & SELEMPBASINFO_RS("gender") & "'"
                a = CInt(SELDTMTAB_RS("hbsagkey"))
                If Len(Trim(SELDTMTAB_RS("HBSAG"))) > 0 Then
                    CrystalReport1.Formulas(6) = "HBSAG = '" & DECRPT(Trim(SELDTMTAB_RS("hbsag")), a) & "'"
                Else
                    CrystalReport1.Formulas(6) = "HBSAG = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HCVKEY"))
                If Len(Trim(SELDTMTAB_RS("HCV"))) > 0 Then
                    CrystalReport1.Formulas(7) = "HCV = '" & DECRPT(Trim(SELDTMTAB_RS("HCV")), a) & "'"
                Else
                    CrystalReport1.Formulas(7) = "HCV = '" & "" & "'"
                End If
                a = CInt(SELDTMTAB_RS("HIVKEY"))
                If Len(Trim(SELDTMTAB_RS("HIV"))) > 0 Then
                    CrystalReport1.Formulas(8) = "HIV = '" & DECRPT(Trim(SELDTMTAB_RS("HIV")), a) & "'"
                Else
                    CrystalReport1.Formulas(8) = "HIV = '" & "" & "'"
                End If
            
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Destination = crptToPrinter
                'CrystalReport1.Destination = crptToWindow
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                CrystalReport1.Action = 1
                        
                'Set HBSAGKEY to 1 so as to recognize its print had already been taken
                UPD_DTMTAB_QR(0) = selpreqnors("REQNnO")
                UPD_DTMTAB_QR.Execute
                
            Else 'Patient
                        
                str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
                CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & selpreqnors("REQNnO") & "'"
                'CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & selpreqnors("REQNnO") & "' and {DTMVIEW1.REPTTM} >='" & frm_tm & "' and {DTMVIEW1.REPTTM} <='" & to_tm & "'"
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISdtm.RPT"
                CrystalReport1.Formulas(0) = "REQNO = '" & selpreqnors("REQNnO") & "'"
                               
                If Trim(prepath) = "" Then
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(1) = "PREPATH = '" & "Previous REQ No." & Left(prepath, 130) & "'"
                End If
                CrystalReport1.Formulas(2) = "USRID = '" & str & "'"
                
                If SELDTMNO_RS.EOF Then
                    CrystalReport1.Formulas(3) = "dtmno = '" & "" & "'"
                Else
                    CrystalReport1.Formulas(3) = "dtmno = '" & SELDTMNO_RS("dtmno") & "'"
                End If
        
                If Not SELDTMTAB_RS.EOF Then
                    a = CInt(SELDTMTAB_RS("hbsagkey"))
        
                    If Len(Trim(SELDTMTAB_RS("HBSAG"))) > 0 Then
                        CrystalReport1.Formulas(4) = "HBSAG = '" & DECRPT(Trim(SELDTMTAB_RS("hbsag")), 2) & "'"
                    Else
                        CrystalReport1.Formulas(4) = "HBSAG = '" & "" & "'"
                    End If
                    If Len(Trim(SELDTMTAB_RS("HCV"))) > 0 Then
                        CrystalReport1.Formulas(5) = "HCV = '" & SELDTMTAB_RS("HCV") & "'"
                    Else
                        CrystalReport1.Formulas(5) = "HCV = '" & "" & "'"
                    End If
                    If Len(Trim(SELDTMTAB_RS("HIV"))) > 0 Then
                        CrystalReport1.Formulas(6) = "HIV = '" & SELDTMTAB_RS("HIV") & "'"
                    Else
                        CrystalReport1.Formulas(6) = "HIV = '" & "" & "'"
                    End If
                End If
    
                '******
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.Destination = crptToPrinter
                'CrystalReport1.Destination = crptToWindow
                CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
                'CrystalReport1.Action = 1
                
                abc = CrystalReport1.PrintReport
                
                  Dim ANS
                  'ANS = MsgBox("GENERATED PRINT OUT IS CORRECT" & vbCrLf & "IF YES THEN CLICK ON 'YES' OTHERWISE CLICK ON 'NO'", vbYesNo, "CONFIRM")
                  'If ANS = vbYes Then
                'Set HBSAGKEY to 1 so as to recognize its print had already been taken
                UPD_DTMTAB_QR(0) = selpreqnors("REQNnO")
                UPD_DTMTAB_QR.Execute
                'End If
       
            End If
            
        Else
            'PRINTOUT ALREADY TAKEN
            p_count = p_count + 1
        End If
        
        selpreqnors.MoveNext
    Wend
    
    
End If

If p_count = 0 Then
    LBL_PRINTED.Caption = ""
    MsgBox "Printing is completed." & vbCrLf & "Total " & selpreqnors.RowCount & " reports printed."
Else
    LBL_PRINTED.Caption = ""
    MsgBox "Printing is completed." & vbCrLf & "Total " & selpreqnors.RowCount - p_count & " reports printed out of " & selpreqnors.RowCount & "," & vbCrLf & "since " & p_count & " printouts were already taken."
End If
lbl_cap.Caption = ""
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_ref_Click()
MSK_FROMDT1.Text = ""
MSK_TODT1.Text = ""
Label1.Caption = ""
LBL_PRINTED.Caption = ""
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
msk_fromdt.Text = ""
msk_todt.Text = ""

Label1.Caption = ""
LBL_PRINTED.Caption = ""
Me.Hide
DIS_MAIN.show
End Sub

Private Sub Form_Load()
Dim str As String

'****************** Shraddha Pandit 18/03/09 ****************
'str = "SELECT REQNNO, CASENO FROM RISLIB.RREQUISITN WHERE MODALITYCD='SS' AND (FREPORTDT BETWEEN ? AND ?)  AND REQNSTAT='T' order by REFBYDOC"
'ABOVE STATEMENT COMMENTED and next added by Shradha on 21/07/09
'str = "SELECT REQNNO, CASENO FROM RISLIB.RREQUISITN WHERE MODALITYCD='SS' AND (FREPORTDT BETWEEN ? AND ?)  AND REQNSTAT='T' order by CASENO"
'above query original 31/08/09
'str = "SELECT reqnno,A.caseno,repttm from rislib.rrequisitn A, dislib.dtmview1 B Where A.reqnno = B.reqno and B.repttm>=? and B.repttm<=? and  A.freportdt Between ? and ?"
'str = "SELECT reqnno,A.caseno,repttm from rislib.rrequisitn A, dislib.dtmview1 B Where A.reqnno = B.reqno and (B.repttm>=? and B.repttm<=?) and  A.Preportdt Between ? and ?"
'Commented and added by Shraddha on 20/04/10
str = "SELECT reqnno, A.caseno, repttm from rislib.rrequisitn A, dislib.dtmview1 B Where A.reqnno = B.reqno and (B.repttm>=? and B.repttm<=?) and  A.Freportdt Between ? and ?"
Set selpreqnoqr = rdb.CreateQuery("", str)

'************************************************************

'****************** Shraddha 31/08/09 ***********************
'str = "SELECT * FROM DISLIB.DTMVIEW1 WHERE (FREPORTDT BETWEEN ? AND ?)  AND (REPTTM >= ? AND REPTTM <= ?)"
'Set sel_repttm_qr = rdb.CreateQuery("", str)

'str = "SELECT reqnno,A.caseno,repttm from rislib.rrequisitn A, dislib.dtmview1 B Where A.reqnno = B.reqno and B.repttm>=? and B.repttm<=? and  A.freportdt Between ? and ?"
'Set sel_repttm_qr = rdb.CreateQuery("", str)
'************************************************************
''SELECT QUERY FOR DTMTAB TABLE
str = "SELECT * FROM DISLIB.DTMTAB WHERE REQNO=?"
Set SELDTMTAB_QR = rdb.CreateQuery("", str)

''SELECT QUERY FOR EMPBASINFO TABLE
str = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
Set SELEMPBASINFO_QR = rdb.CreateQuery("SELEMPBASINFOQR", str)

''SELECT QUERY FOR RREQUISITN TABLE
str = "select * from rislib.rrequisitn where reqnno=?"
Set req2_qr = rdb.CreateQuery("", str)

''SELECT QUERY FOR DTMNO TABLE
str = "select * from dislib.dtmno where caseno=?"
Set SELDTMNO_QR = rdb.CreateQuery("", str)

''SELECT QUERY FOR DTMVIEW TABLE
str = "select * from dislib.DREQUISITN where REQNO =?"
Set DrequisitnQR = rdb.CreateQuery("", str)

str = "UPDATE DISLIB.DTMTAB SET HBSAGKEY=1 WHERE REQNO=? "
Set UPD_DTMTAB_QR = rdb.CreateQuery("UPDREQUISITN_QR1", str)
    
datetime

End Sub

Public Function DECRPT(INPSTR As String, KEYVAL As Integer)
Dim ts(1 To 80) As String
Dim OUTSTR As String
Dim m, n As Integer
Dim tempcahr
m = 1
n = Len(Trim(INPSTR))
tempcahr = Trim(INPSTR)

While n > 0
    ts(m) = Chr(Asc(Mid(tempcahr, m, 1)) - KEYVAL - m)
    m = m + 1
    n = n - 1
Wend

n = 1
m = m - 1

While m > 0
    OUTSTR = OUTSTR + ts(n)
    n = n + 1
    m = m - 1
Wend

DECRPT = OUTSTR
End Function

Private Sub MSK_FROMDT1_GotFocus()
MSK_FROMDT1.Text = Format(DATE11, "dd/mm/yyyy")
End Sub

Private Sub MSK_TODT1_GotFocus()
MSK_TODT1.Text = Format(DATE11, "dd/mm/yyyy")
End Sub
