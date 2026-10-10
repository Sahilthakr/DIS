VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_ClinicalPharmacologyRCV 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Clinical Pharmacology - Sample Receiving"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14580
   LinkTopic       =   "Form2"
   ScaleHeight     =   10020
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   855
      Left            =   3960
      TabIndex        =   8
      Top             =   1080
      Width           =   6015
      Begin VB.TextBox TXTREQNO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2040
         MaxLength       =   13
         TabIndex        =   10
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisition No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1650
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
      Height          =   3615
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   14535
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   16761024
         ForeColorFixed  =   128
         GridColor       =   0
         AllowUserResizing=   3
         FormatString    =   $"frm_ClinicalPharmacologyRCV.frx":0000
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
      BackColor       =   &H00E0E0E0&
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
      Height          =   975
      Left            =   3120
      TabIndex        =   0
      Top             =   6000
      Width           =   5415
      Begin VB.CommandButton cmd_back 
         BackColor       =   &H00FFC0C0&
         Caption         =   "BACK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2040
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Back to the Main Screen"
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton CMD_REFRESH 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Refresh Grid"
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3600
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Print Report"
         Top             =   240
         Width           =   1200
      End
   End
   Begin MSWinsockLib.Winsock WIN_CLIENT 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblorg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "HOMI BHABHA CANCER HOSPITAL AND RESEARCH CENTRE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1080
      TabIndex        =   11
      Top             =   0
      Width           =   11955
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "SAMPLE RECEIVING - CLINICAL PHARMACOLOGY"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   480
      Width           =   8595
   End
   Begin VB.Label lblDisp 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      Top             =   7200
      Width           =   3855
   End
End
Attribute VB_Name = "frm_ClinicalPharmacologyRCV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Prabhu Yeshu
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
Dim REQNOB As String
Dim REQNOBARCODE As String
Dim REQ111 As String
Dim PAT_NAME As String
Dim RCVDATE As String
Dim RCVTIME As String

Dim MODALITYB As String
Dim DEPTB As String
Dim LEN1 As Integer
Dim str As String

Dim sel_rtimelog_qr As rdoQuery
Dim sel_rtimelog_rs As rdoResultset

Private Sub cmd_back_Click()
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_exit_Click()
End
End Sub

Private Sub cmd_refresh_Click()
'Image1.Picture = LoadPicture("")
MSFlexGrid1.clear
MSFlexGrid1.Rows = 2

'*****
MSFlexGrid1.ColWidth(0) = 600
MSFlexGrid1.ColWidth(1) = 2000
MSFlexGrid1.ColWidth(2) = 1300
MSFlexGrid1.ColWidth(3) = 2700
MSFlexGrid1.ColWidth(4) = 1600
MSFlexGrid1.ColWidth(5) = 1600
MSFlexGrid1.ColWidth(6) = 1600
MSFlexGrid1.ColWidth(7) = 1600

MSFlexGrid1.TextMatrix(0, 0) = "SrNo."
MSFlexGrid1.TextMatrix(0, 1) = "Requisition No."
MSFlexGrid1.TextMatrix(0, 2) = "CaseNo."
MSFlexGrid1.TextMatrix(0, 3) = "Name"
MSFlexGrid1.TextMatrix(0, 4) = "Collection Date"
MSFlexGrid1.TextMatrix(0, 5) = "Collection Time"
MSFlexGrid1.TextMatrix(0, 6) = "Received Date"
MSFlexGrid1.TextMatrix(0, 7) = "Received Time"
TXTREQNO.Text = ""

MSFlexGrid1.Rows = 2
'*****

TXTREQNO.SetFocus
End Sub

Private Sub Form_Activate()
TXTREQNO.SetFocus
End Sub

Private Sub Form_Load()
 lblorg.Caption = G_OrgFullName
 If G_UnitCode = "KM" Then
    frm_ClinicalPharmacologyRCV.BackColor = RGB(255, 255, 192)
 ElseIf G_UnitCode = "KR" Then
    frm_ClinicalPharmacologyRCV.BackColor = RGB(255, 216, 177)
 End If
str = "SELECT DEPT_CODE,MODLITYDES FROM RISLIB.RMODALITY WHERE MODALITYCD=?"
Set SELDEPT_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set GETCASE_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RTIMELOG WHERE REQNNO=?"
Set sel_rtimelog_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM PATOOLSLIB.TA101 WHERE CASE_NO=?"
Set GETNAME_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=?"
Set GETNAMESTAFF_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELREQNSTAT_QR = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RTIMELOG SET SAMRECDATE=CURRENT DATE, SAMRECTIME=CURRENT TIME WHERE REQNNO=? "
Set UPDRTIME_QR = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT=?,EXAMINATN=? WHERE REQNNO=?"
Set UPDRREQ_QR = rdb.CreateQuery("", str)

DATETIME

End Sub


Private Sub TXTREQNO_Change()
If Len(Trim(TXTREQNO.Text)) = 11 Then

    lblDisp.Caption = ""
    DoEvents
    READDATA
    
End If
End Sub

Public Sub READDATA()

'On Error GoTo MYERR

    DATETIME
    RCVDATE = DATE11
    RCVTIME = time1
    
    REQNOB = Trim(TXTREQNO.Text)
    
    REQ111 = Trim(Mid(REQNOB, 3, 8))
    
    MODALITYB = Trim(UCase(Left(REQNOB, 2)))
        
    REQNOBARCODE = MODALITYB & REQ111
        
    SELDEPT_QR(0) = MODALITYB
    Set SELDEPT_RS = SELDEPT_QR.OpenResultset(1, 3)
    DEPTB = Trim(SELDEPT_RS("DEPT_CODE"))
          
    ''get reqno
    REQNOB = DEPTB & REQNOBARCODE
    
    SELREQNSTAT_QR(0) = REQNOB
    Set SELREQNSTAT_RS = SELREQNSTAT_QR.OpenResultset(1, 3)
    
    Dim ST1 As String
    ST1 = "C"
    If Trim(SELREQNSTAT_RS("REQNSTAT")) = "M" Then
        DoEvents
        MsgBox "Payment not done", vbOKOnly
    ElseIf Trim(SELREQNSTAT_RS("REQNSTAT")) = "Z" Then
        DoEvents
        MsgBox "Refunded", vbOKOnly
    ElseIf Trim(SELREQNSTAT_RS("REQNSTAT")) = "P" Then
        DoEvents
        MsgBox "Payment done but sample not yet collected", vbOKOnly
        DoEvents
    ElseIf Trim(SELREQNSTAT_RS("REQNSTAT")) = "R" Then
        DoEvents
        MsgBox "Sample marked as Repeat", vbOKOnly
        DoEvents
    ElseIf Trim(SELREQNSTAT_RS("REQNSTAT")) = "E" Then
        DoEvents
        MsgBox "Sample already recieved at department", vbOKOnly
        DoEvents
    ElseIf Trim(SELREQNSTAT_RS("REQNSTAT")) = "T" Then
        DoEvents
        MsgBox "Preliminary report done", vbOKOnly
        DoEvents
    ElseIf Trim(SELREQNSTAT_RS("REQNSTAT")) = "F" Then
        DoEvents
        MsgBox "Final report done", vbOKOnly
        DoEvents
    ElseIf Trim(SELREQNSTAT_RS("REQNSTAT")) = "C" Then
    
        PAT_NAME = ""
    
        GETCASE_QR(0) = REQNOB
        Set GETCASE_RS = GETCASE_QR.OpenResultset(1, 3)
    
        sel_rtimelog_qr(0) = REQNOB
        Set sel_rtimelog_rs = sel_rtimelog_qr.OpenResultset(1, 3)
     
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
            .Rows = .Rows + 1
            .row = .Rows - 2
            
                .TextMatrix(.row, 0) = .row
                .TextMatrix(.row, 1) = REQNOB
                .TextMatrix(.row, 2) = GETCASE_RS("CASENO")
                .TextMatrix(.row, 3) = PAT_NAME
                .TextMatrix(.row, 4) = sel_rtimelog_rs("SAMCOLDATE")
                .TextMatrix(.row, 5) = sel_rtimelog_rs("SAMCOLTIME")
                .TextMatrix(.row, 6) = RCVDATE
                .TextMatrix(.row, 7) = RCVTIME
                
            End With
            rdb.BeginTrans
                assign
            rdb.CommitTrans
            
    Else
          MsgBox "Please Check the query Status"
    End If
            
    DoEvents
    TXTREQNO.Text = ""
    TXTREQNO.SetFocus
Exit Sub

MYERR:
        rdb.RollbackTrans
        MsgBox (Err.Description)
    
End Sub

Public Sub assign()
'UPDATE RISLIB.RTIMELOG SET SAMRECDATE=CURRENT DATE, SAMRECTIME=CURRENT TIME WHERE REQNNO=?
    UPDRTIME_QR(0) = REQNOB
    UPDRTIME_QR.Execute
'UPDATE RISLIB.RREQUISITN SET REQNSTAT=?,LOCATION=? WHERE REQNNO=?
    UPDRREQ_QR(0) = "E"
    UPDRREQ_QR(1) = WIN_CLIENT.LocalIP
    UPDRREQ_QR(2) = REQNOB
    UPDRREQ_QR.Execute
        
End Sub

