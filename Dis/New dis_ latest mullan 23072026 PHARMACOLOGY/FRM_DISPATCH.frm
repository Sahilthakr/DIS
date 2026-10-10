VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FRM_DISPATCH 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "REPORTS DISPATCH FOR WARDS"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7665
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6240
      Top             =   4200
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   6135
      Begin VB.CommandButton CMD_EXIT 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   5
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton CMD_CLEAR 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton CMD_REPORT 
         Caption         =   "&Generate Report"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   2520
         Width           =   2055
      End
      Begin MSMask.MaskEdBox MSK_FTIME 
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MSK_TTIME 
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MSK_TDATE 
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   720
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MSK_FDATE 
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   720
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "ENTER TIME"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "ENTER DATE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FROM"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Left            =   720
         TabIndex        =   15
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Left            =   3240
         TabIndex        =   14
         Top             =   720
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FROM"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Left            =   720
         TabIndex        =   13
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Left            =   3240
         TabIndex        =   12
         Top             =   1800
         Width           =   270
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0FF&
         Height          =   735
         Left            =   480
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   5175
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0FF&
         Height          =   735
         Left            =   480
         Shape           =   4  'Rounded Rectangle
         Top             =   1560
         Width           =   5175
      End
   End
   Begin VB.PictureBox P1 
      Height          =   255
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   4320
      Width           =   1935
   End
   Begin VB.PictureBox SB1 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   7605
      TabIndex        =   1
      Top             =   4815
      Width           =   7665
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   2760
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\reports\Yearrep_qty.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "BIOCHEMISTRY REPORTS DISPATCH FOR WARDS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   840
      TabIndex        =   16
      Top             =   120
      Width           =   6180
   End
End
Attribute VB_Name = "FRM_DISPATCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ipADDR
Dim PROCESS
Dim MsgTitle

Dim SEL_BIODISPCH_rq As rdoQuery
Dim SEL_BIODISPCH_RS As rdoResultset

Dim InsBIOCH_WARD_rq As rdoQuery
Dim UpdBIOCH_WARD_rq As rdoQuery

'WARD WISE
Dim SEL_WARD_XII_RQ As rdoQuery
Dim SEL_WARD_XII_RS As rdoResultset
Dim SEL_WARD_CR_RQ As rdoQuery
Dim SEL_WARD_CR_RS As rdoResultset
Dim SEL_WARD_X_RQ As rdoQuery
Dim SEL_WARD_X_RS As rdoResultset
Dim SEL_WARD_IX_RQ As rdoQuery
Dim SEL_WARD_IX_RS As rdoResultset
Dim SEL_WARD_RMC_RQ As rdoQuery
Dim SEL_WARD_RMC_RS As rdoResultset
Dim SEL_WARD_MALE_RQ As rdoQuery
Dim SEL_WARD_MALE_RS As rdoResultset
Dim SEL_WARD_FEMALE_RQ As rdoQuery
Dim SEL_WARD_FEMALE_RS As rdoResultset
Dim SEL_WARD_SPW_RQ As rdoQuery
Dim SEL_WARD_SPW_RS As rdoResultset
Dim SEL_WARD_PW_RQ As rdoQuery
Dim SEL_WARD_PW_RS As rdoResultset
Dim SEL_WARD_DLX_RQ As rdoQuery
Dim SEL_WARD_DLX_RS As rdoResultset
Dim SEL_WARD_ICU_RQ As rdoQuery
Dim SEL_WARD_ICU_RS As rdoResultset
Dim SEL_WARD_RECOV_RQ As rdoQuery
Dim SEL_WARD_RECOV_RS As rdoResultset



Dim str As String
Dim typ As Integer

Private Sub cmd_clear_Click()
'Unload Me
SETVALUES
End Sub

Private Sub cmd_exit_Click()
End
End Sub

Private Sub CMD_REPORT_Click()

Me.MousePointer = 11

If MSK_FDATE <> "__/__/____" Then
         If Not ValidateDate(MSK_FDATE) Then
            MSK_FDATE.Text = "__/__/____"
            MSK_FDATE.SetFocus
            Me.MousePointer = 0
            Exit Sub
            
         End If
            'MSK_TDATE.SetFocus
Else
        MsgBox "From Date can not be Empty", vbCritical, MsgTitle
        MSK_FDATE.SetFocus
        Me.MousePointer = 0
        Exit Sub
        
End If

If MSK_TDATE <> "__/__/____" Then
         If Not ValidateDate(MSK_TDATE) Then
            MSK_TDATE.Text = "__/__/____"
            MSK_TDATE.SetFocus
            Me.MousePointer = 0
            Exit Sub
            
         End If
Else
        MsgBox "To Date can not be Empty", vbCritical, MsgTitle
            MSK_TDATE.SetFocus
        Me.MousePointer = 0
        Exit Sub
        
End If
    
If DateDiff("d", MSK_FDATE, MSK_TDATE) < 0 Then
        MsgBox "Todate can not be Less than From date", vbCritical, MsgTitle
        MSK_TDATE.Text = "__/__/____"
        MSK_TDATE.SetFocus
        Me.MousePointer = 0
        Exit Sub
End If

'' *****************************TIME

If MSK_FTIME <> "__:__:__" Then
         If Not ValidateTIME(MSK_FTIME) Then
            MSK_FTIME.Text = "__:__:__"
            MSK_FTIME.SetFocus
            Me.MousePointer = 0
            Exit Sub
            
         End If
            'MSK_TTIME.SetFocus
Else
        MsgBox "From time can not be Empty", vbCritical, MsgTitle
        MSK_FTIME.SetFocus
        Me.MousePointer = 0
        Exit Sub
        
End If

If MSK_TTIME <> "__:__:__" Then
         If Not ValidateTIME(MSK_TTIME) Then
            MSK_TTIME.Text = "__:__:__"
            MSK_TTIME.SetFocus
            Me.MousePointer = 0
            Exit Sub
            
         End If
Else
        MsgBox "To time can not be Empty", vbCritical, MsgTitle
            MSK_TTIME.SetFocus
        Me.MousePointer = 0
        Exit Sub
        
End If
    
If DateDiff("S", MSK_FTIME, MSK_TTIME) < 0 Then
        MsgBox "To time can not be Less than From time", vbCritical, MsgTitle
        MSK_TTIME.Text = "__:__:__"
        MSK_TTIME.SetFocus
        Me.MousePointer = 0
        Exit Sub
End If
P1.Visible = True
''''''TEMP TABLE
time1 = Time
PROCESS = "WAIT...."
STATUSB
TEMPTBL
STATUSB
PRINTRPT
STATUSB
time2 = Time
PROCESS = "COMPLETED!!!!"
STATUSB
MousePointer = vbNormal
P1.Visible = False
STATUSB
'MsgBox "total time taken : " & DateDiff("n", time1, time2)
End Sub
Sub PRINTRPT()
'CR1.ReportFileName = App.Path & "\" & "\REPORTS\WARDWISE.rpt"
CR1.ReportFileName = "C:\DISREPORTSEC\WARDWISE.rpt"
CR1.Destination = crptToWindow
CR1.Formulas(0) = "FDATE='" & Trim(MSK_FDATE.Text) & "'"
CR1.Formulas(1) = "TDATE='" & Trim(MSK_TDATE.Text) & "'"
CR1.Formulas(2) = "FTIME='" & Trim(MSK_FTIME.Text) & "'"
CR1.Formulas(3) = "TTIME='" & Trim(MSK_TTIME.Text) & "'"

CR1.Formulas(4) = "ip ='" & Winsock1.LocalIP & "'"
'CR1.Formulas(5) = "uid='" & strlogin & "'"

CR1.WindowState = crptMaximized

CR1.Action = 1

P1.Value = P1.Value + 5
STATUSB
End Sub

Sub SETQUERIES()
''strsql = "SELECT WARD_NAME,COUNT(CASENO) AS TC FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
''strsql = strsql + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "'"
''strsql = strsql + " GROUP BY WARD_NAME ORDER BY TC DESC"
''CNT_BIODISPCH_rq.Close
''Set CNT_BIODISPCH_rq = RdCn.CreateQuery("InsBIOCH_WARD", strsql)
''CNT_BIODISPCH_rq.QueryTimeout = 500
''Set CNT_BIODISPCH_RS = CNT_BIODISPCH_rq.OpenResultset(1, 1)
''
'''If CNT_BIODISPCH_RS.EOF = False Then
'''CNT_BIODISPCH_RS
'''End If
''While Not CNT_BIODISPCH_RS.EOF
''AB = CNT_BIODISPCH_RS("TC")
''CNT_BIODISPCH_RS.MoveNext
''
''Wend
STRSQL = "SELECT * FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
STRSQL = STRSQL + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "' AND WARD_NAME = 'ANNEXE IX FLOOR WARD'"
STRSQL = STRSQL + " ORDER BY PATHNO,CASENO"
'SEL_BIODISPCH_rq.Close
Set SEL_BIODISPCH_rq = RdCn.CreateQuery("InsBIOCH_WARD", STRSQL)
SEL_BIODISPCH_rq.QueryTimeout = 500
Set SEL_BIODISPCH_RS = SEL_BIODISPCH_rq.OpenResultset(1, 1)

STRSQL = "Insert into DISLIB.BIOCH_WARD (AB_XII,AB_XI,AB_X,AB_IX,AB_RMC,MB_MALE,MB_FEMALE,MB_SPW,MB_PVT,MB_DLX,MB_ICU,MB_RECOV)"
STRSQL = STRSQL + " values (?,?,?,?,?,?,?,?,?,?,?,?)"
Set InsBIOCH_WARD_rq = RdCn.CreateQuery("", STRSQL)

End Sub

Sub SETQUERIES_ALL()
STATUSB
'ANNEXE XII FLOOR WARD
STRSQL = ""
STRSQL = "SELECT * FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
STRSQL = STRSQL + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "' AND (TRIM(WARD_NAME) = 'ANNEXE XII FLOOR WAR' OR TRIM(WARD_NAME) = 'CHEMOTHERAPY WARD') "
STRSQL = STRSQL + " ORDER BY PATHNO,CASENO"
'SEL_WARD_XII_RQ.Close
Set SEL_WARD_XII_RQ = RdCn.CreateQuery("InsBIOCH_WARD1", STRSQL)
SEL_WARD_XII_RQ.QueryTimeout = 300
Set SEL_WARD_XII_RS = SEL_WARD_XII_RQ.OpenResultset(1, 1)


'ANNEXE CLINICAL RES. WARD
STRSQL = ""
STRSQL = "SELECT * FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
STRSQL = STRSQL + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "' AND TRIM(WARD_NAME) = 'CLINICAL RES WARD' "
STRSQL = STRSQL + " ORDER BY PATHNO,CASENO"
Set SEL_WARD_CR_RQ = RdCn.CreateQuery("InsBIOCH_WARD2", STRSQL)
SEL_WARD_CR_RQ.QueryTimeout = 300
Set SEL_WARD_CR_RS = SEL_WARD_CR_RQ.OpenResultset(1, 1)
STATUSB
'ANNEXE X FLOOR WARD
STRSQL = ""
STRSQL = "SELECT * FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
STRSQL = STRSQL + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "' AND TRIM(WARD_NAME) = 'ANNEXE X FLOOR WARD'  "
STRSQL = STRSQL + " ORDER BY PATHNO,CASENO"
Set SEL_WARD_X_RQ = RdCn.CreateQuery("InsBIOCH_WARD3", STRSQL)
SEL_WARD_X_RQ.QueryTimeout = 300
Set SEL_WARD_X_RS = SEL_WARD_X_RQ.OpenResultset(1, 1)

'ANNEXE IX FLOOR WARD
STRSQL = ""
STRSQL = "SELECT * FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
STRSQL = STRSQL + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "' AND TRIM(WARD_NAME) = 'ANNEXE IX FLOOR WARD' "
STRSQL = STRSQL + " ORDER BY PATHNO,CASENO"
Set SEL_WARD_IX_RQ = RdCn.CreateQuery("InsBIOCH_WARD4", STRSQL)
SEL_WARD_IX_RQ.QueryTimeout = 300
Set SEL_WARD_IX_RS = SEL_WARD_IX_RQ.OpenResultset(1, 1)

'ANNEXE RMC  WARD
STRSQL = ""
STRSQL = "SELECT * FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
STRSQL = STRSQL + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "' AND TRIM(WARD_NAME) = 'RMC WARD' "
STRSQL = STRSQL + " ORDER BY PATHNO,CASENO"
Set SEL_WARD_RMC_RQ = RdCn.CreateQuery("InsBIOCH_WARD5", STRSQL)
SEL_WARD_RMC_RQ.QueryTimeout = 300
Set SEL_WARD_RMC_RS = SEL_WARD_RMC_RQ.OpenResultset(1, 1)
STATUSB
'MAIN MALE  WARD
STRSQL = ""
STRSQL = "SELECT * FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
STRSQL = STRSQL + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "' AND TRIM(WARD_NAME) = 'MALE WARD'"
STRSQL = STRSQL + " ORDER BY PATHNO,CASENO"
Set SEL_WARD_MALE_RQ = RdCn.CreateQuery("InsBIOCH_WARD6", STRSQL)
SEL_WARD_MALE_RQ.QueryTimeout = 300
Set SEL_WARD_MALE_RS = SEL_WARD_MALE_RQ.OpenResultset(1, 1)

'MAIN FEMALE  WARD
STRSQL = ""
STRSQL = "SELECT * FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
STRSQL = STRSQL + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "' AND TRIM(WARD_NAME) = 'FEMALE WARD' "
STRSQL = STRSQL + " ORDER BY PATHNO,CASENO"
Set SEL_WARD_FEMALE_RQ = RdCn.CreateQuery("InsBIOCH_WARD7", STRSQL)
SEL_WARD_FEMALE_RQ.QueryTimeout = 300
Set SEL_WARD_FEMALE_RS = SEL_WARD_FEMALE_RQ.OpenResultset(1, 1)
STATUSB
'MAIN SPW  WARD
STRSQL = ""
STRSQL = "SELECT * FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
STRSQL = STRSQL + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "' AND TRIM(WARD_NAME) = 'SEMI PRIVATE WARD' "
STRSQL = STRSQL + " ORDER BY PATHNO,CASENO"
Set SEL_WARD_SPW_RQ = RdCn.CreateQuery("InsBIOCH_WARD8", STRSQL)
SEL_WARD_SPW_RQ.QueryTimeout = 300
Set SEL_WARD_SPW_RS = SEL_WARD_SPW_RQ.OpenResultset(1, 1)

'MAIN PW  WARD
STRSQL = ""
STRSQL = "SELECT * FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
STRSQL = STRSQL + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "' AND TRIM(WARD_NAME) = 'PRIVATE WARD' "
STRSQL = STRSQL + " ORDER BY PATHNO,CASENO"
Set SEL_WARD_PW_RQ = RdCn.CreateQuery("InsBIOCH_WARD9", STRSQL)
SEL_WARD_PW_RQ.QueryTimeout = 300
Set SEL_WARD_PW_RS = SEL_WARD_PW_RQ.OpenResultset(1, 1)

'MAIN DLX  WARD
STRSQL = ""
STRSQL = "SELECT * FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
STRSQL = STRSQL + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "' AND (TRIM(WARD_NAME) = 'DELUX WARD' OR TRIM(WARD_NAME) = 'SPECIAL WARD')"
STRSQL = STRSQL + " ORDER BY PATHNO,CASENO"
'SEL_WARD_DLX_RQ.Close
Set SEL_WARD_DLX_RQ = RdCn.CreateQuery("InsBIOCH_WARD10", STRSQL)
SEL_WARD_DLX_RQ.QueryTimeout = 300
Set SEL_WARD_DLX_RS = SEL_WARD_DLX_RQ.OpenResultset(1, 1)

'MAIN ICU  WARD
STRSQL = ""
STRSQL = "SELECT * FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
STRSQL = STRSQL + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "' AND ICDS2 = 'YES'"
STRSQL = STRSQL + " ORDER BY PATHNO,CASENO"
Set SEL_WARD_ICU_RQ = RdCn.CreateQuery("InsBIOCH_WARD11", STRSQL)
SEL_WARD_ICU_RQ.QueryTimeout = 300
Set SEL_WARD_ICU_RS = SEL_WARD_ICU_RQ.OpenResultset(1, 1)

'MAIN RECOV  WARD
STRSQL = ""
STRSQL = "SELECT * FROM DISLIB.BIODISPCH WHERE PREPORTDT BETWEEN '" & Format(Trim(MSK_FDATE.Text), "YYYY-MM-DD") & "' AND '" & Format(Trim(MSK_TDATE.Text), "YYYY-MM-DD") & "'"
STRSQL = STRSQL + " AND SAVREPTIME BETWEEN '" & Trim(MSK_FTIME.Text) & "' AND '" & Trim(MSK_TTIME.Text) & "' AND ICDS2 = 'YR'"
STRSQL = STRSQL + " ORDER BY PATHNO,CASENO"
Set SEL_WARD_RECOV_RQ = RdCn.CreateQuery("InsBIOCH_WARD12", STRSQL)
SEL_WARD_RECOV_RQ.QueryTimeout = 300
Set SEL_WARD_RECOV_RS = SEL_WARD_RECOV_RQ.OpenResultset(1, 1)

STRSQL = ""
STRSQL = "Insert into DISLIB.BIOCH_WARD (AB_XII,AB_XI,AB_X,AB_IX,AB_RMC,MB_MALE,MB_FEMALE,MB_SPW,MB_PVT,MB_DLX,MB_ICU,MB_RECOV,FLD1)"
STRSQL = STRSQL + " values (?,?,?,?,?,?,?,?,?,?,?,?,?)"
Set InsBIOCH_WARD_rq = RdCn.CreateQuery("", STRSQL)
STATUSB
End Sub


Sub TEMPTBL()
DEL_SQL = "DELETE FROM DISLIB.BIOCH_WARD"
RdCn.Execute DEL_SQL
P1.Value = P1.Value + 2
INSERTDATA
End Sub
Function GETNUMROWS() As Double
Dim ROW_NUM As Double

If SEL_WARD_XII_RS.RowCount <= SEL_WARD_CR_RS.RowCount Then
ROW_NUM = SEL_WARD_CR_RS.RowCount
Else
ROW_NUM = SEL_WARD_XII_RS.RowCount
End If

If ROW_NUM < SEL_WARD_X_RS.RowCount Then
ROW_NUM = SEL_WARD_X_RS.RowCount
End If

If ROW_NUM < SEL_WARD_IX_RS.RowCount Then
ROW_NUM = SEL_WARD_IX_RS.RowCount
End If

If ROW_NUM < SEL_WARD_RMC_RS.RowCount Then
ROW_NUM = SEL_WARD_RMC_RS.RowCount
End If

If ROW_NUM < SEL_WARD_MALE_RS.RowCount Then
ROW_NUM = SEL_WARD_MALE_RS.RowCount
End If

If ROW_NUM < SEL_WARD_FEMALE_RS.RowCount Then
ROW_NUM = SEL_WARD_FEMALE_RS.RowCount
End If

If ROW_NUM < SEL_WARD_FEMALE_RS.RowCount Then
ROW_NUM = SEL_WARD_FEMALE_RS.RowCount
End If

If ROW_NUM < SEL_WARD_SPW_RS.RowCount Then
ROW_NUM = SEL_WARD_SPW_RS.RowCount
End If

If ROW_NUM < SEL_WARD_DLX_RS.RowCount Then
ROW_NUM = SEL_WARD_DLX_RS.RowCount
End If

If ROW_NUM < SEL_WARD_ICU_RS.RowCount Then
ROW_NUM = SEL_WARD_ICU_RS.RowCount
End If

If ROW_NUM < SEL_WARD_RECOV_RS.RowCount Then
ROW_NUM = SEL_WARD_RECOV_RS.RowCount
End If
GETNUMROWS = ROW_NUM
End Function
Sub INSERTDATA()
Dim i As Double
P1.Value = P1.Value + 2
STATUSB
SETQUERIES_ALL
P1.Value = P1.Value + 5
STATUSB
Rows1 = GETNUMROWS
P1.Value = P1.Value + 1
STATUSB
For i = 1 To Rows1
If SEL_WARD_XII_RS.EOF = False Then
    If ((Trim(SEL_WARD_XII_RS("ICDS2")) <> "YES" Or Trim(SEL_WARD_XII_RS("ICDS2")) <> "YR") And Not (Trim(SEL_WARD_XII_RS("ICDS2")) = Null)) Then
    AB_XII = ""
    Else
    AB_XII = Trim(SEL_WARD_XII_RS("PATHNO")) & "-" & Trim(SEL_WARD_XII_RS("CASENO"))
    End If
Else
AB_XII = ""
End If

If SEL_WARD_CR_RS.EOF = False Then
    If ((Trim(SEL_WARD_CR_RS("ICDS2")) <> "YES" Or Trim(SEL_WARD_CR_RS("ICDS2")) <> "YR") And Not (Trim(SEL_WARD_CR_RS("ICDS2")) = Null)) Then
    AB_XI = ""
    Else
    AB_XI = Trim(SEL_WARD_CR_RS("PATHNO")) & "-" & Trim(SEL_WARD_CR_RS("CASENO"))
    End If
Else
AB_XI = ""
End If

If SEL_WARD_X_RS.EOF = False Then
    If (Trim(SEL_WARD_X_RS("ICDS2")) <> "YES" Or Trim(SEL_WARD_X_RS("ICDS2")) <> "YR") Then
    AB_X = ""
    Else
    AB_X = Trim(SEL_WARD_X_RS("PATHNO")) & "-" & Trim(SEL_WARD_X_RS("CASENO"))
    End If
Else
AB_X = ""
End If

If SEL_WARD_IX_RS.EOF = False Then
    If ((Trim(SEL_WARD_IX_RS("ICDS2")) <> "YES" Or Trim(SEL_WARD_IX_RS("ICDS2")) <> "YR") And Not (Trim(SEL_WARD_IX_RS("ICDS2")) = Null)) Then
    AB_IX = ""
    Else
    AB_IX = Trim(SEL_WARD_IX_RS("PATHNO")) & "-" & Trim(SEL_WARD_IX_RS("CASENO"))
    End If
Else
AB_IX = ""
End If

If SEL_WARD_RMC_RS.EOF = False Then
    If ((Trim(SEL_WARD_RMC_RS("ICDS2")) <> "YES" Or Trim(SEL_WARD_RMC_RS("ICDS2")) <> "YR") And Not (Trim(SEL_WARD_RMC_RS("ICDS2")) = Null)) Then
    AB_RMC = ""
    Else
    AB_RMC = Trim(SEL_WARD_RMC_RS("PATHNO")) & "-" & Trim(SEL_WARD_RMC_RS("CASENO"))
    End If
Else
AB_RMC = ""
End If

If SEL_WARD_MALE_RS.EOF = False Then
    If ((Trim(SEL_WARD_MALE_RS("ICDS2")) <> "YES" Or Trim(SEL_WARD_MALE_RS("ICDS2")) <> "YR") And Not (Trim(SEL_WARD_MALE_RS("ICDS2")) = Null)) Then
    MB_MALE = ""
    Else
    MB_MALE = Trim(SEL_WARD_MALE_RS("PATHNO")) & "-" & Trim(SEL_WARD_MALE_RS("CASENO"))
    End If
Else
MB_MALE = ""
End If

If SEL_WARD_FEMALE_RS.EOF = False Then
    If ((Trim(SEL_WARD_FEMALE_RS("ICDS2")) <> "YES" Or Trim(SEL_WARD_FEMALE_RS("ICDS2")) <> "YR") And Not (Trim(SEL_WARD_FEMALE_RS("ICDS2")) = Null)) Then
    MB_FEMALE = ""
    Else
    MB_FEMALE = Trim(SEL_WARD_FEMALE_RS("PATHNO")) & "-" & Trim(SEL_WARD_FEMALE_RS("CASENO"))
    End If

Else
MB_FEMALE = ""
End If

If SEL_WARD_SPW_RS.EOF = False Then
    If ((Trim(SEL_WARD_SPW_RS("ICDS2")) <> "YES" Or Trim(SEL_WARD_SPW_RS("ICDS2")) <> "YR") And Not (Trim(SEL_WARD_SPW_RS("ICDS2")) = Null)) Then
    MB_SPW = ""
    Else
    MB_SPW = Trim(SEL_WARD_SPW_RS("PATHNO")) & "-" & Trim(SEL_WARD_SPW_RS("CASENO"))
    End If
Else
MB_SPW = ""
End If
 

If SEL_WARD_PW_RS.EOF = False Then
    If ((Trim(SEL_WARD_PW_RS("ICDS2")) <> "YES" Or Trim(SEL_WARD_PW_RS("ICDS2")) <> "YR") And Not (Trim(SEL_WARD_PW_RS("ICDS2")) = Null)) Then
    MB_PW = ""
    Else
    MB_PW = Trim(SEL_WARD_PW_RS("PATHNO")) & "-" & Trim(SEL_WARD_PW_RS("CASENO"))
    End If
Else
MB_PW = ""
End If

If SEL_WARD_DLX_RS.EOF = False Then
    If ((Trim(SEL_WARD_DLX_RS("ICDS2")) <> "YES" Or Trim(SEL_WARD_DLX_RS("ICDS2")) <> "YR") And Not (Trim(SEL_WARD_DLX_RS("ICDS2")) = Null)) Then
    MB_DLX = ""
    Else
    MB_DLX = Trim(SEL_WARD_DLX_RS("PATHNO")) & "-" & Trim(SEL_WARD_DLX_RS("CASENO"))
    End If
Else
MB_DLX = ""
End If

If SEL_WARD_ICU_RS.EOF = False Then
    If ((Trim(SEL_WARD_ICU_RS("ICDS2")) <> "YES" Or Trim(SEL_WARD_ICU_RS("ICDS2")) <> "YR") And Not (Trim(SEL_WARD_ICU_RS("ICDS2")) = Null)) Then
    MB_ICU = ""
    Else
    MB_ICU = Trim(SEL_WARD_ICU_RS("PATHNO")) & "-" & Trim(SEL_WARD_ICU_RS("CASENO"))
    End If
Else
MB_ICU = ""
End If

If SEL_WARD_RECOV_RS.EOF = False Then
    If ((Trim(SEL_WARD_RECOV_RS("ICDS2")) <> "YES" Or Trim(SEL_WARD_RECOV_RS("ICDS2")) <> "YR") And Not (Trim(SEL_WARD_RECOV_RS("ICDS2")) = Null)) Then
    MB_RECOV = ""
    Else
    MB_RECOV = Trim(SEL_WARD_RECOV_RS("PATHNO")) & "-" & Trim(SEL_WARD_RECOV_RS("CASENO"))
    End If
Else
MB_RECOV = ""
End If

       InsBIOCH_WARD_rq(0) = AB_XII
       InsBIOCH_WARD_rq(1) = AB_XI
       InsBIOCH_WARD_rq(2) = AB_X
       InsBIOCH_WARD_rq(3) = AB_IX
       InsBIOCH_WARD_rq(4) = AB_RMC
       InsBIOCH_WARD_rq(5) = MB_MALE
       InsBIOCH_WARD_rq(6) = MB_FEMALE
       InsBIOCH_WARD_rq(7) = MB_SPW
       InsBIOCH_WARD_rq(8) = MB_PW
       InsBIOCH_WARD_rq(9) = MB_DLX
       InsBIOCH_WARD_rq(10) = MB_ICU
       InsBIOCH_WARD_rq(11) = MB_RECOV
       InsBIOCH_WARD_rq(12) = i
       InsBIOCH_WARD_rq.Execute
       
On Error Resume Next
SEL_WARD_XII_RS.MoveNext
SEL_WARD_CR_RS.MoveNext
SEL_WARD_X_RS.MoveNext
SEL_WARD_IX_RS.MoveNext
SEL_WARD_RMC_RS.MoveNext
SEL_WARD_MALE_RS.MoveNext
SEL_WARD_FEMALE_RS.MoveNext
SEL_WARD_SPW_RS.MoveNext
SEL_WARD_DLX_RS.MoveNext
SEL_WARD_ICU_RS.MoveNext
SEL_WARD_RECOV_RS.MoveNext
SEL_WARD_PW_RS.MoveNext
STATUSB
PERC = (85 * i) / Rows1
P1.Value = P1.Value + PERC
Next
 
End Sub
Private Sub Form_Load()
P1.Visible = False
Main1
ipADDR = Winsock1.LocalIP
STATUSB
SETVALUES
MsgTitle = "BIOCHEMISTRY REPORTS DISPATCH FOR WARD"
End Sub

Sub SETVALUES()
MSK_FDATE.Text = Format(Date, "DD/MM/YYYY")
MSK_TDATE.Text = Format(Date, "DD/MM/YYYY")
MSK_FTIME.Text = "09:00:00"
MSK_TTIME.Text = Format(Time, "HH:MM:SS")
End Sub

Sub STATUSB()
SB1.Panels(1) = PROCESS
SB1.Panels(2) = ""
SB1.Panels(3) = Format(Date, "DD/MM/YYYY")
SB1.Panels(4) = Format(Time, "HH:MM:SS")
SB1.Panels(5) = ipADDR
End Sub

