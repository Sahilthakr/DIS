VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_CBCworklist 
   BackColor       =   &H00C0C0C0&
   Caption         =   "CBC Worklist"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12270
   LinkTopic       =   "Form2"
   ScaleHeight     =   9060
   ScaleWidth      =   12270
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_back 
      BackColor       =   &H0080C0FF&
      Caption         =   "&BACK"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton cmd_refresh 
      BackColor       =   &H0080C0FF&
      Caption         =   "&REFRESH"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton CMD_EXIT 
      BackColor       =   &H0080C0FF&
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7800
      Width           =   1335
   End
   Begin VB.TextBox TXT_MESSAGE 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1440
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   7320
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.Frame Frame2 
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
      Height          =   5535
      Left            =   480
      TabIndex        =   9
      Top             =   1440
      Width           =   8415
      Begin VB.CheckBox chk_edit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   195
         Left            =   6480
         TabIndex        =   14
         Top             =   840
         Width           =   255
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4575
         Left            =   480
         TabIndex        =   10
         Top             =   480
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   8070
         _Version        =   393216
         Rows            =   20
         Cols            =   5
         FixedCols       =   0
         FormatString    =   "Sr. No.|CASE NO.       | REQUISITION NO.   |REPORT DATE| EDIT"
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
      Begin VB.Shape Shape2 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   5295
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   8175
      End
   End
   Begin VB.CommandButton Command3 
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
      Left            =   2400
      TabIndex        =   4
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&CANCEL"
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
      Left            =   4080
      TabIndex        =   3
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
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
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2280
      Top             =   720
   End
   Begin MSMask.MaskEdBox msk_sysdate 
      Height          =   315
      Left            =   9720
      TabIndex        =   5
      Top             =   360
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
      Left            =   9720
      TabIndex        =   6
      Top             =   720
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
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1680
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MESSAGE"
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
      Left            =   480
      TabIndex        =   13
      Top             =   7440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Developed by Monika and varinder Commence Date : 02/11/21"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1095
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   8
      Top             =   720
      Width           =   675
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.Label LBL_DISW 
      BackStyle       =   0  'Transparent
      Caption         =   "CBC - WORKLIST"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frm_CBCworklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sel_hint_qr As rdoQuery
Dim sel_hint_rs As rdoResultset

Dim sel_inner_hint_qr As rdoQuery
Dim sel_inner_hint_rs As rdoResultset

Dim str As String
Dim var_date As String
Dim ipadd As String

Dim ROUTINE As Boolean
Dim ENDOFPROCESS As Boolean

Dim last15_date

Dim i, I2, K As Long
Dim sel_chkrreq_qr As rdoQuery
Dim sel_chkrreq_rs As rdoResultset

Dim sel_dhem_qr As rdoQuery
Dim sel_dhem_rs As rdoResultset

Dim panic_v As Boolean

Private Sub CHK_EDIT_Click()
If chk_edit.Value = 1 Then
    'frm_CBCedit.show
    
    frm_CBCedit.cmd_commit.Caption = "COMMIT"
    frm_CBCedit.cmd_commit.BackColor = &H80000003
    frm_CBCedit.cmd_commit.Enabled = True
    frm_CBCedit.cmd_commit.Visible = True
    
    CBCrec = False
    'panic = False
    Call frm_CBCedit.fillData(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1), MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))
    If CBCrec = False Then
        frm_CBCedit.show
    Else
        frm_CBCedit.Hide
        frm_CBCworklist.show
    End If
ElseIf chk_edit.Value = 0 Then
    frm_CBCedit.Hide
End If
End Sub

Private Sub cmd_back_Click()
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_exit_Click()
ENDOFPROCESS = True
End
End Sub

Public Sub cmd_refresh_Click()

frm_CBCedit.cmd_commit.Caption = "COMMIT"
frm_CBCedit.cmd_commit.BackColor = &H80000003

'cmd_commit.Caption = "COMMIT"
'cmd_commit.BackColor = &H80000003


MSFlexGrid1.clear
chk_edit.Visible = False
chk_edit.Value = 0
MSFlexGrid1.FormatString = "Sr. No.|CASE NO.       | REQUISITION NO.   |REPORT DATE| EDIT"
DoEvents
FILLGRID
End Sub

Private Sub Form_Load()

'str = "select distinct(caseno), reqno, TDATE, HB,PLATELETS from interface.hinterface where TDATE >=? and IFLAG='2' and MACHINENO='LH750' order by caseno"
'str = "select distinct(caseno), reqno, TDATE, HB,PLATELETS from interface.hinterface where TDATE >=? and IFLAG='2' and MACHINENO='SAPP2' order by caseno"
'str = "select distinct(caseno), reqno, TDATE, HB,PLATELETS from interface.hinterface where TDATE >=? and IFLAG='2' order by caseno"
'Condition of machine name ommitted on 01/12/09 by Shraddha
'Condition of MACHINENO='LH5750' added on 26/11/09 by Shraddha
'Condition of MACHINENO='SAPP2' added on 03/12/09 by Shraddha
'str = "select distinct(caseno), reqno, TDATE, HB,PLATELETS from interface.hinterface where TDATE >=? and IFLAG='2' and MACHINENO='SAPP2' order by caseno"
str = "select distinct(caseno), reqno, TDATE, HB,PLATELETS from interface.hinterface where TDATE >=? and IFLAG='2' and MACHINENO='LH750' order by caseno"
Set sel_hint_qr = rdb.CreateQuery("", str)

'----- Shraddha Pandit 23/11/09 -----
str = "select distinct(REQNNO),freportdt,REMARKS from rislib.rrequisitn A, DISLIB.DHEMCBCDET B where A.REQNNO=B.REQNO AND CASENO=? and MODALITYCD='HL' and FREPORTDT>=? order by  reqnno desc fetch first 3 rows only"
Set sel_chkrreq_qr = rdb.CreateQuery("", str)
'------------------------------------
'***** Shradha Pandit 23/11/09 *****
'Bring panic value records which are still hanging in the system
'str = "select distinct(caseno), reqno, TDATE from interface.hinterface where TDATE >=? and IFLAG='2' and (HB<'4' or PLATELETS<'5000') order by caseno"
'If sel_HangRec_rs.EOF And Trim(selhintfrs("HB")) > 4 And Trim(selhintfrs("PLATELETS")) > 5000 Then
'End If
str = "select * from DISLIB.DHEMCBCDET where REQNO =?"
Set sel_dhem_qr = rdb.CreateQuery("", str)
'***********************************
datetime
    
CBCrec = False
'panic = False
panic_v = False
ipadd = Winsock1.LocalIP

msk_sysdate.Text = Format(DATE11, "DD/MM/YYYY")
msk_systime.Text = Format(Time, "HH:MM:SS")
End Sub

Public Sub FILLGRID()

MSFlexGrid1.Rows = 1
i = 1
    
'TXT_MESSAGE.Text = "  RETRIEVING DETAILS PLEASE WAIT......"

last15_date = DateAdd("d", -15, msk_sysdate)
last15_date = Format(last15_date, "yyyy-mm-dd")
var_date = Year(last15_date) & "," & month(last15_date) & "," & Day(last15_date)

sel_hint_qr(0) = var_date
Set sel_hint_rs = sel_hint_qr.OpenResultset(1, 3)

If Not sel_hint_rs.EOF Then
    sel_hint_rs.MoveFirst

    While sel_hint_rs.EOF = False
        
        '***** Shraddha Pandit 232/11/09 *****
        'str = "select DISTINCT(REQNNO),freportdt,REMARKS from rislib.rrequisitn A, DISLIB.DHEMCBCDET B where A.REQNNO=B.REQNO AND CASENO=? and MODALITYCD='HL' and FREPORTDT>=? order by  reqnno desc fetch first 3 rows only"
        'Set sel_chkrreq_qr = rdb.CreateQuery("", str)
        sel_chkrreq_qr(0) = sel_hint_rs("caseno")
        sel_chkrreq_qr(1) = var_date
        Set sel_chkrreq_rs = sel_chkrreq_qr.OpenResultset(1, 3)
                
        If sel_chkrreq_rs.EOF Then
            'sel_hint_rs.MoveNext
            
            '###########################################
            
            '+++++
                'Panic value check
                
                
                    If Trim(sel_hint_rs("HB")) < 4 Or Trim(sel_hint_rs("PLATELETS")) < 5 Then 'Beep
                        'panic_v = True
                        For S = 1 To 10
                            Beep
                        Next S
                        MSFlexGrid1.AddItem i & vbTab & sel_hint_rs("CASENO") & vbTab & req_chglen(Trim(sel_hint_rs("reqno"))) & vbTab & sel_hint_rs("TDATE")
                        i = i + 1
                    End If
                    
  
            '+++++
            sel_hint_rs.MoveNext
            
            'panic_v = False
            '###########################################
            
            
        Else
            If sel_chkrreq_rs.RowCount < 3 Then
                '+++++
                'Panic value check
                sel_dhem_qr(0) = sel_chkrreq_rs("REQNNO")
                Set sel_dhem_rs = sel_dhem_qr.OpenResultset(1, 3)
                
                For a = 1 To sel_dhem_rs.RowCount
                    If Trim(sel_dhem_rs("PRO_CODE")) = "HAEMOGLOBIN" Then
                        If Trim(sel_dhem_rs("OBSERVALUE")) < 4 Then 'Beep
                            panic_v = True
                            For S = 1 To 10
                                Beep
                            Next S
                        
                            MSFlexGrid1.AddItem i & vbTab & sel_hint_rs("CASENO") & vbTab & req_chglen(Trim(sel_hint_rs("reqno"))) & vbTab & sel_hint_rs("TDATE")
                            i = i + 1
                        End If
                    End If
                    
                    If Trim(sel_dhem_rs("PRO_CODE")) = "PLATELETS" Then
                        If Trim(sel_dhem_rs("OBSERVALUE")) < 5 Then 'Beep
                            panic_v = True
                            For S = 1 To 10
                                Beep
                            Next S
                        
                            MSFlexGrid1.AddItem i & vbTab & sel_hint_rs("CASENO") & vbTab & req_chglen(Trim(sel_hint_rs("reqno"))) & vbTab & sel_hint_rs("TDATE")
                            i = i + 1
                        End If
                    End If
                    
                    sel_dhem_rs.MoveNext
              
                Next
                
                '$$$$$$$$$$$$$$$$$$$$$$$$$$$$
                If panic_v = False Then
                    MSFlexGrid1.AddItem i & vbTab & sel_hint_rs("CASENO") & vbTab & req_chglen(Trim(sel_hint_rs("reqno"))) & vbTab & sel_hint_rs("TDATE")
                    i = i + 1
                    sel_hint_rs.MoveNext
                Else
                    sel_dhem_rs.MoveNext
                End If
                '$$$$$$$$$$$$$$$$$$$$$$$$$$$$
                
                 'sel_hint_rs.MoveNext
            '+++++
            Else
            'If sel_chkrreq_rs.RowCount >= 3 Then
        '*************************************
                'Next 3 are original statements
                MSFlexGrid1.AddItem i & vbTab & sel_hint_rs("CASENO") & vbTab & req_chglen(Trim(sel_hint_rs("reqno"))) & vbTab & sel_hint_rs("TDATE")
                i = i + 1
                sel_hint_rs.MoveNext
        '**********************************
            End If
        End If
        '******** Shraddha 23/11/09 *******
        panic_v = False
    Wend

End If
Exit Sub

End Sub


Private Sub Timer1_Timer()
    datetime
    msk_sysdate.Text = Format(DATE11, "DD/MM/YYYY")
    msk_systime.Text = Format(time1, "HH:MM:SS")
End Sub


Private Sub chk_edit_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If MSFlexGrid1.row = MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
    End If
    If MSFlexGrid1.row < MSFlexGrid1.Rows - 1 Then
        MSFlexGrid1.row = MSFlexGrid1.row + 1
        chk_edit = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
        chk_edit.SetFocus
        chk_edit.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
    End If
End If
Exit Sub

End Sub

Private Sub MSFlexGrid1_Click()
chk_edit.Visible = True
chk_edit = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 4))
chk_edit.SetFocus
chk_edit.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
Exit Sub
End Sub

