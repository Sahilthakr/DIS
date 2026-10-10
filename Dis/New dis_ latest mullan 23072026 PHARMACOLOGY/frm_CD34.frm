VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_CD34 
   BackColor       =   &H00C0C0C0&
   Caption         =   "CD34 Enumeration"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12045
   LinkTopic       =   "Form2"
   ScaleHeight     =   8715
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_commit 
      Caption         =   "CO&MMIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   26
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox txt_reqno 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   25
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txt_caseno 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   24
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   9000
      TabIndex        =   19
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   9000
      TabIndex        =   18
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmd_print 
      Caption         =   "&PRINT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   17
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmd_SAVE 
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmd_modify 
      Caption         =   "&MODIFY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmd_clear 
      Caption         =   "&CLEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   15
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TXT_ABScd34CellCount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      MaxLength       =   8
      TabIndex        =   4
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox TXT_CD34cellcount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      MaxLength       =   8
      TabIndex        =   3
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox TXT_ManualMNC 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox TXT_TTLMNC 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      MaxLength       =   8
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox TXT_TTLWBCcount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmd_Back 
      Caption         =   "&BACK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Top             =   6120
      Width           =   975
   End
   Begin Crystal.CrystalReport CR_CD34 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "cells/uL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   29
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "x 10log3/uL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   28
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Completed on 03/11/21"
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   6375
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Requisition No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Case No. "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   22
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   8520
      TabIndex        =   21
      Top             =   480
      Width           =   360
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   8520
      TabIndex        =   20
      Top             =   840
      Width           =   420
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      Height          =   3375
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   6375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Absolute CD34 cell count"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "CD34 cell count % "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Manual MNC %    (Lymphocyte + Monocyte)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   12
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total MNC % (Automated)  (Lymphocyte + Monocyte)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total WBC Count (Automated)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000013&
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
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   0
      Width           =   9255
   End
   Begin VB.Label lbl_frmCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "CD34 Enumeration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
End
Attribute VB_Name = "frm_CD34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim sql1 As String
Dim sql2 As String
Dim sql3 As String
Dim sql4 As String
Dim sql5 As String
Dim sql6 As String
Dim sql7 As String
Dim str As String

Dim SEL_HCD34_QR As rdoQuery
Dim SEL_HCD34_RS As rdoResultset

'***** Shraddha Pandit 22/09/09 *****
Dim INS_HCD34_QR As rdoQuery
Dim UPD_HCD34_QR As rdoQuery
Dim sel_rtimelog_qr As rdoQuery
Dim sel_rtimelog_rs As rdoResultset

Dim upd_rrequisitn_qr As rdoQuery
Dim updrrequisitn_qr As rdoQuery
Dim upd_rtimelog_qr As rdoQuery
Dim upd_commit_CD34 As rdoQuery

Dim sel_rrequisitn_qr As rdoQuery
Dim sel_rrequisitn_rs As rdoResultset
'************************************

Private Sub Form_Load()
datetime
txt_date.Text = Format(DATE11, "DD/MM/YYYY")
txt_time.Text = time1

txt_caseno.Text = hemcbcreport.MaskEdBox1.FormattedText
txt_reqno = req_chglen(hemcbcreport.cmb_reqno)

sql = "SELECT * FROM DISLIB.HAEMATCD34 WHERE REQNO=?"
Set SEL_HCD34_QR = rdb.CreateQuery("", sql)
 
sql1 = "INSERT INTO dislib.haematcd34 (CASENO,REQNO,PATNAME,SEX,AGE,CATG,STATUS,WARD," & _
"BEDNO,HEMNO,PRIMEDOC,REFDOC,PROVDIAGNO,SAVEDATE,SAVETIME,TTLWBCCNT,UOMTTLWBC,TTLMNCP," & _
"MANUALMNCP,CD34CELCNT,ABS34CELCT,UOMABSCD34,USERID,IPADDR,REQNDATE,REQNTIME,SAMCOLDATE," & _
"SAMCOLTIME,SAMRECDATE,SAMRECTIME,COMREPDATE,COMREPTIME,FRMCAPTION) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
Set INS_HCD34_QR = rdb.CreateQuery("", sql1)
'13 9 10
'CASENO REQNO PATNAME SEX AGE CATG STATUS WARD BEDNO HEMNO PRIMEDOC REFDOC PROVDIAGNO
'SAVEDATE SAVETIME TTLWBCCNT UOMTTLWBC TTLMNCP MANUALMNCP CD34CELCNT ABS34CELCT UOMABSCD34
'USERID IPADDR REQNDATE REQNTIME SAMCOLDATE SAMCOLTIME SAMRECDATE SAMRECTIME COMREPDATE COMREPTIME

'*** OGI '7 9 8 9 4
'sql1 = "INSERT INTO dislib.haematcd34 (CASENO,REQNO,PATNAME,SEX,AGE,CATG,STATUS," & _
'"WARD,BEDNO,HEMNO,PRIMEDOC,REFDOC,PROVDIAGNO,SAVEDATE,SAVETIME,WBCDILUCNT," & _
'"DLUFACTOR,UNDLUWBCCT,MNCCTDLU,MNCCTUNDLU,MANUMNCPER,G4EVENTS,G1EVENTS,CD34CELCT," & _
'"ABS34CELCT,CD3CELCT,ABS3CELCT,USERID,IPADDR,REQNDATE,REQNTIME,SAMCOLDATE,SAMCOLTIME," & _
'"SAMRECDATE,SAMRECTIME,COMREPDATE,COMREPTIME) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
'Set INS_HCD34_QR = rdb.CreateQuery("", sql1)
'*******
'14
sql2 = "UPDATE dislib.haematcd34 SET TTLWBCCNT=?,TTLMNCP=?,MANUALMNCP=?,CD34CELCNT=?,ABS34CELCT=?," & _
"REQNDATE=?,REQNTIME=?,SAMCOLDATE=?,SAMCOLTIME=?,SAMRECDATE=?,SAMRECTIME=?,COMREPDATE=?,COMREPTIME=? WHERE REQNO=?"

'5 8 8
'*** Ogi
'sql2 = "UPDATE dislib.haematcd34 SET WBCDILUCNT=?,DLUFACTOR=?,UNDLUWBCCT=?,MNCCTDLU=?,MNCCTUNDLU=?," & _
'"MANUMNCPER=?,G4EVENTS=?,G1EVENTS=?,CD34CELCT=?,ABS34CELCT=?,CD3CELCT=?,ABS3CELCT=?, REQNDATE=?," & _
'"REQNTIME=?,SAMCOLDATE=?,SAMCOLTIME=?,SAMRECDATE=?,SAMRECTIME=?,COMREPDATE=?,COMREPTIME=? WHERE REQNO=?"
'***

Set UPD_HCD34_QR = rdb.CreateQuery("", sql2)

'30/11/09
str = "update Rislib.rrequisitn set ENTERBY=?, PREPORTDT=?,PREPORTBY=?,REQNSTAT=?,FREPORTBY=?,FREPORTDT=? WHERE REQNNO=?"
Set updrrequisitn_qr = rdb.CreateQuery("", str)
   

sql3 = "SELECT * from rislib.rtimelog where reqnno=?"
Set sel_rtimelog_qr = rdb.CreateQuery("", sql3)

sql4 = "UPDATE RISLIB.RREQUISITN SET FREPORTBY=?, FREPORTDT=?, REQNSTAT=?, COMMITFLAG=? WHERE REQNNO=?"
Set upd_rrequisitn_qr = rdb.CreateQuery("", sql4)

sql5 = "UPDATE RISLIB.RTIMELOG SET COMREPDATE=?,COMREPTIME=? WHERE REQNNO=?"
Set upd_rtimelog_qr = rdb.CreateQuery("", sql5)

sql6 = "UPDATE dislib.haematcd34 SET REQNDATE=?,REQNTIME=?,SAMCOLDATE=?,SAMCOLTIME=?,SAMRECDATE=?,SAMRECTIME=?,COMREPDATE=?,COMREPTIME=? WHERE REQNO=?"
Set upd_commit_CD34 = rdb.CreateQuery("", sql6)

sql7 = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set sel_rrequisitn_qr = rdb.CreateQuery("", sql7)

If SURFACEFORM.opt_PrePBSC_34 = True Then
    lbl_frmCaption.Caption = "Pre PBSC CD34 Enumeration"
ElseIf SURFACEFORM.opt_PBSCharvest_34 = True Then
    lbl_frmCaption.Caption = "PBSC Harvest CD34 Enumeration"
ElseIf SURFACEFORM.opt_BMharvest_34 = True Then
    lbl_frmCaption.Caption = "BM Harvest CD34 Enumeration"
End If

If hemcbcreport.txt_NEW.Text = True Then
    'MsgBox "NEW"

    'Button level validation
    cmd_save.Enabled = True
    cmd_back.Enabled = True
    'cmd_clear.Enabled = True
    
    cmd_modify.Enabled = False
    cmd_commit.Enabled = False
    cmd_print.Enabled = False
    'cmd_Exit.Enabled = False

    updFlfCD34 = False
    insFlgCD34 = True
    
    TXT_TTLWBCcount.Text = ""
    TXT_TTLMNC.Text = ""
    TXT_ManualMNC.Text = ""
    TXT_CD34cellcount.Text = ""
    TXT_ABScd34CellCount.Text = ""

    TXT_TTLWBCcount.Enabled = True
    TXT_TTLMNC.Enabled = True
    TXT_ManualMNC.Enabled = True
    TXT_CD34cellcount.Enabled = True
    TXT_ABScd34CellCount.Enabled = True

ElseIf hemcbcreport.txt_VIEW.Text = True Or hemcbcreport.txt_SAVE = True Then
    'MsgBox "VIEW"
    'Button level validation
    cmd_modify.Enabled = True
    'cmd_clear.Enabled = True
    cmd_print.Enabled = True
    cmd_back.Enabled = True
    
    cmd_save.Enabled = False
    cmd_commit.Enabled = False
    'cmd_Exit.Enabled = False
    'Check for commit **********

    sel_rrequisitn_qr(0) = req_chglen(hemcbcreport.cmb_reqno)
    Set sel_rrequisitn_rs = sel_rrequisitn_qr.OpenResultset(1, 3)

    If sel_rrequisitn_rs.RowCount > 0 Then 'if commited
        If sel_rrequisitn_rs("COMMITFLAG") = "Y" Then
            MsgBox "Final Report is Generated, You can't Modify the Final Report"
            cmd_modify.Enabled = False
            cmd_commit.Enabled = False
            cmd_save.Enabled = False
            cmd_print.Enabled = True
        End If
    Else
    '********** If not commited *********

        

        updFlfCD34 = True
        insFlgCD34 = False

    End If 'Commit check
    sel_rrequisitn_rs.Close

        SEL_HCD34_QR(0) = req_chglen(hemcbcreport.cmb_reqno)
        Set SEL_HCD34_RS = SEL_HCD34_QR.OpenResultset(1, 3)

        If Not SEL_HCD34_RS.EOF Then

            TXT_TTLWBCcount = SEL_HCD34_RS("TTLWBCCNT")
            TXT_TTLMNC = SEL_HCD34_RS("TTLMNCP")
            TXT_ManualMNC = SEL_HCD34_RS("MANUALMNCP")
            TXT_CD34cellcount = SEL_HCD34_RS("CD34CELCNT")
            TXT_ABScd34CellCount = SEL_HCD34_RS("ABS34CELCT")
            
        End If

        TXT_TTLWBCcount.Enabled = False
        TXT_TTLMNC.Enabled = False
        TXT_ManualMNC.Enabled = False
        TXT_CD34cellcount.Enabled = False
        TXT_ABScd34CellCount.Enabled = False

ElseIf hemcbcreport.txt_SAVE.Text = "1" Then
    'MsgBox "Insert"
ElseIf hemcbcreport.txt_SAVE.Text = "2" Then
    'MsgBox "Update"
End If
End Sub

Private Sub cmd_clear_Click()
    TXT_TTLWBCcount.Text = ""
    TXT_TTLMNC.Text = ""
    TXT_ManualMNC.Text = ""
    TXT_CD34cellcount.Text = ""
    TXT_ABScd34CellCount.Text = ""
        
    cmd_save.Enabled = False
    cmd_print.Enabled = False
End Sub

Private Sub cmd_modify_Click()
updFlfCD34 = True
insFlgCD34 = False

'Button level validation
cmd_save.Enabled = True
cmd_back.Enabled = True
'cmd_clear.Enabled = True

cmd_modify.Enabled = False
cmd_commit.Enabled = False
cmd_print.Enabled = False
'cmd_Exit.Enabled = False

TXT_TTLWBCcount.Enabled = True
TXT_TTLMNC.Enabled = True
TXT_ManualMNC.Enabled = True
TXT_CD34cellcount.Enabled = True
TXT_ABScd34CellCount.Enabled = True

End Sub

Private Sub cmd_save_Click()
sel_rtimelog_qr(0) = req_chglen(hemcbcreport.cmb_reqno.Text)
Set sel_rtimelog_rs = sel_rtimelog_qr.OpenResultset(1, 3)

If TXT_TTLWBCcount.Text = "" Then
    MsgBox "Please enter Total WBC Count (Automated)"
    TXT_TTLWBCcount.SetFocus
    Exit Sub
    
ElseIf TXT_TTLMNC.Text = "" Then
    MsgBox "Please enter Total MNC % (Automated)"
    TXT_TTLMNC.SetFocus
    Exit Sub
    
ElseIf TXT_ManualMNC.Text = "" Then
    MsgBox "Please enter Manual MNC %"
    TXT_ManualMNC.SetFocus
    Exit Sub

ElseIf TXT_CD34cellcount.Text = "" Then
    MsgBox "Please enter CD34 cell count %"
    TXT_CD34cellcount.SetFocus
    Exit Sub

ElseIf TXT_ABScd34CellCount.Text = "" Then
    MsgBox "Please enter Absolute CD34 cell count"
    TXT_ABScd34CellCount.SetFocus
    Exit Sub

End If

If insFlgCD34 = True Then
    INS_HCD34_QR(32) = Trim(lbl_frmCaption.Caption)
    INS_HCD34_QR(0) = hemcbcreport.MaskEdBox1.FormattedText
    INS_HCD34_QR(1) = req_chglen(Trim(hemcbcreport.cmb_reqno.Text))
    INS_HCD34_QR(2) = Trim(hemcbcreport.txt_name.Text)
    INS_HCD34_QR(3) = Trim(hemcbcreport.txt_sex.Text)
    INS_HCD34_QR(4) = Trim(hemcbcreport.txt_Age.Text)
    INS_HCD34_QR(5) = Trim(hemcbcreport.txt_catEgory.Text)
    INS_HCD34_QR(6) = Trim(hemcbcreport.txt_status.Text)
    INS_HCD34_QR(7) = Trim(hemcbcreport.TXT_WARD.Text)
    INS_HCD34_QR(8) = Trim(hemcbcreport.TXT_BEDNO.Text)
    INS_HCD34_QR(9) = Trim(hemcbcreport.TXT_PATH.Text)
    INS_HCD34_QR(10) = Trim(hemcbcreport.txt_pridoc.Text)

    INS_HCD34_QR(11) = Trim(hemcbcreport.txt_refdoc.Text)
    INS_HCD34_QR(12) = Trim(hemcbcreport.txt_procedure.Text)
    INS_HCD34_QR(13) = Format(txt_date, "YYYY-MM-DD")
    INS_HCD34_QR(14) = Format(txt_time, "HH:MM:SS")
    '13 9 10
'CASENO REQNO PATNAME SEX AGE CATG STATUS WARD BEDNO HEMNO PRIMEDOC REFDOC PROVDIAGNO
'SAVEDATE SAVETIME TTLWBCCNT UOMTTLWBC TTLMNCP MANUALMNCP CD34CELCNT ABS34CELCT UOMABSCD34
'USERID IPADDR REQNDATE REQNTIME SAMCOLDATE SAMCOLTIME SAMRECDATE SAMRECTIME COMREPDATE COMREPTIME
    INS_HCD34_QR(15) = Trim(TXT_TTLWBCcount.Text)
    INS_HCD34_QR(16) = "10log3/uL"
    
    INS_HCD34_QR(17) = Trim(TXT_TTLMNC.Text)
    INS_HCD34_QR(18) = Trim(TXT_ManualMNC.Text)
    INS_HCD34_QR(19) = Trim(TXT_CD34cellcount.Text)
    
    INS_HCD34_QR(20) = Trim(TXT_ABScd34CellCount.Text)
    INS_HCD34_QR(21) = "cells/ul"
    
    INS_HCD34_QR(22) = Right(LOGIN, 6) 'UID
    INS_HCD34_QR(23) = Trim(ipadd) 'IPADDR

    If Not sel_rtimelog_rs.EOF Then
        If sel_rtimelog_rs("REQNDATE") <> "" Then
            INS_HCD34_QR(24) = sel_rtimelog_rs("REQNDATE")
        Else
            INS_HCD34_QR(24) = Null
        End If
        If sel_rtimelog_rs("REQNTIME") <> "" Then
            INS_HCD34_QR(25) = sel_rtimelog_rs("REQNTIME")
        Else
            INS_HCD34_QR(25) = Null
        End If
        If sel_rtimelog_rs("SAMCOLDATE") <> "" Then
            INS_HCD34_QR(26) = sel_rtimelog_rs("SAMCOLDATE")
        Else
            INS_HCD34_QR(26) = Null
        End If
        If sel_rtimelog_rs("SAMCOLTIME") <> "" Then
            INS_HCD34_QR(27) = sel_rtimelog_rs("SAMCOLTIME")
        Else
            INS_HCD34_QR(27) = Null
        End If
        If sel_rtimelog_rs("SAMRECDATE") <> "" Then
            INS_HCD34_QR(28) = sel_rtimelog_rs("SAMRECDATE")
        Else
            INS_HCD34_QR(28) = Null
        End If
        If sel_rtimelog_rs("SAMRECTIME") <> "" Then
            INS_HCD34_QR(29) = sel_rtimelog_rs("SAMRECTIME")
        Else
            INS_HCD34_QR(29) = Null
        End If
        If sel_rtimelog_rs("COMREPDATE") <> "" Then
            INS_HCD34_QR(30) = sel_rtimelog_rs("COMREPDATE")
        Else
            INS_HCD34_QR(30) = Null
        End If
        If sel_rtimelog_rs("COMREPTIME") <> "" Then
            INS_HCD34_QR(31) = sel_rtimelog_rs("COMREPTIME")
        Else
            INS_HCD34_QR(31) = Null
        End If

    End If
    
    INS_HCD34_QR.Execute
    '***** Shraddha 30/11/09 *****
    updrrequisitn_qr(0) = hemcbcreport.cmb_enterby.Text
    updrrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(2) = hemcbcreport.regcbo.Text ' cmb_finalby.Text
    updrrequisitn_qr(3) = "T"
    updrrequisitn_qr(4) = hemcbcreport.cmb_finalby.Text
    updrrequisitn_qr(5) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    updrrequisitn_qr(6) = Trim(req_chglen(hemcbcreport.cmb_reqno.Text))
    updrrequisitn_qr.Execute
    '*****************************
    MsgBox "Record inserted successfully"

End If

If updFlfCD34 = True Then

'14
'sql2 = "UPDATE dislib.haematcd34 SET TTLWBCCNT=?,TTLMNCP=?,MANUALMNCP=?,CD34CELCNT=?,ABS34CELCT=?," & _
'"REQNDATE=?,REQNTIME=?,SAMCOLDATE=?,SAMCOLTIME=?,SAMRECDATE=?,SAMRECTIME=?,COMREPDATE=?,COMREPTIME=? WHERE REQNO=?"

    UPD_HCD34_QR(0) = Trim(TXT_TTLWBCcount.Text)
    UPD_HCD34_QR(1) = Trim(TXT_TTLMNC.Text)
    UPD_HCD34_QR(2) = Trim(TXT_ManualMNC.Text)
    UPD_HCD34_QR(3) = Trim(TXT_CD34cellcount.Text)
    UPD_HCD34_QR(4) = Trim(TXT_ABScd34CellCount.Text)
    
    If Not sel_rtimelog_rs.EOF Then
        If sel_rtimelog_rs("REQNDATE") <> "" Then
            UPD_HCD34_QR(5) = sel_rtimelog_rs("REQNDATE")
        Else
            UPD_HCD34_QR(5) = Null
        End If
        If sel_rtimelog_rs("REQNTIME") <> "" Then
            UPD_HCD34_QR(6) = sel_rtimelog_rs("REQNTIME")
        Else
            UPD_HCD34_QR(6) = Null
        End If
        If sel_rtimelog_rs("SAMCOLDATE") <> "" Then
            UPD_HCD34_QR(7) = sel_rtimelog_rs("SAMCOLDATE")
        Else
            UPD_HCD34_QR(7) = Null
        End If
        If sel_rtimelog_rs("SAMCOLTIME") <> "" Then
            UPD_HCD34_QR(8) = sel_rtimelog_rs("SAMCOLTIME")
        Else
            UPD_HCD34_QR(8) = Null
        End If
        If sel_rtimelog_rs("SAMRECDATE") <> "" Then
            UPD_HCD34_QR(9) = sel_rtimelog_rs("SAMRECDATE")
        Else
            UPD_HCD34_QR(9) = Null
        End If
        If sel_rtimelog_rs("SAMRECTIME") <> "" Then
            UPD_HCD34_QR(10) = sel_rtimelog_rs("SAMRECTIME")
        Else
            UPD_HCD34_QR(10) = Null
        End If
        If sel_rtimelog_rs("COMREPDATE") <> "" Then
            UPD_HCD34_QR(11) = sel_rtimelog_rs("COMREPDATE")
        Else
            UPD_HCD34_QR(11) = Null
        End If
        If sel_rtimelog_rs("COMREPTIME") <> "" Then
            UPD_HCD34_QR(12) = sel_rtimelog_rs("COMREPTIME")
        Else
            UPD_HCD34_QR(12) = Null
        End If
        UPD_HCD34_QR(13) = req_chglen(Trim(hemcbcreport.cmb_reqno.Text))
    End If

    UPD_HCD34_QR.Execute

    MsgBox "Record updated successfully"
End If

sel_rtimelog_rs.Close

cmd_save.Enabled = False
cmd_modify.Enabled = False
'cmd_clear.Enabled = False
cmd_commit.Enabled = True
cmd_print.Enabled = True
cmd_back.Enabled = True
'*********************** 22/09/09 **********************

End Sub

Private Sub cmd_commit_Click()

'*************** 24/09/09 *************

 rdb.BeginTrans
    upd_rrequisitn_qr(0) = UCase(hemcbcreport.cmb_finalby.Text)
    upd_rrequisitn_qr(1) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
    upd_rrequisitn_qr(2) = "F"
    upd_rrequisitn_qr(3) = "Y"
    upd_rrequisitn_qr(4) = req_chglen(Trim(hemcbcreport.cmb_reqno.Text))

    upd_rrequisitn_qr.Execute

    upd_rtimelog_qr(0) = Format(txt_date.Text, "YYYY-MM-DD")
    upd_rtimelog_qr(1) = Format(txt_time.Text, "HH:MM:SS")
    upd_rtimelog_qr(2) = req_chglen(Trim(hemcbcreport.cmb_reqno.Text))

    upd_rtimelog_qr.Execute

    MsgBox "Record is Committed Successfully."
rdb.CommitTrans
'**************************************

sel_rtimelog_qr(0) = req_chglen(Trim(hemcbcreport.cmb_reqno.Text))
Set sel_rtimelog_rs = sel_rtimelog_qr.OpenResultset(1, 3)

  If Not sel_rtimelog_rs.EOF Then
        If sel_rtimelog_rs("REQNDATE") <> "" Then
            upd_commit_CD34(0) = sel_rtimelog_rs("REQNDATE")
        Else
            upd_commit_CD34(0) = Null
        End If
        If sel_rtimelog_rs("REQNTIME") <> "" Then
            upd_commit_CD34(1) = sel_rtimelog_rs("REQNTIME")
        Else
            upd_commit_CD34(1) = Null
        End If
        If sel_rtimelog_rs("SAMCOLDATE") <> "" Then
            upd_commit_CD34(2) = sel_rtimelog_rs("SAMCOLDATE")
        Else
            upd_commit_CD34(2) = Null
        End If
        If sel_rtimelog_rs("SAMCOLTIME") <> "" Then
            upd_commit_CD34(3) = sel_rtimelog_rs("SAMCOLTIME")
        Else
            upd_commit_CD34(3) = Null
        End If
        If sel_rtimelog_rs("SAMRECDATE") <> "" Then
            upd_commit_CD34(4) = sel_rtimelog_rs("SAMRECDATE")
        Else
            upd_commit_CD34(4) = Null
        End If
        If sel_rtimelog_rs("SAMRECTIME") <> "" Then
            upd_commit_CD34(5) = sel_rtimelog_rs("SAMRECTIME")
        Else
            upd_commit_CD34(5) = Null
        End If
        If sel_rtimelog_rs("COMREPDATE") <> "" Then
            upd_commit_CD34(6) = sel_rtimelog_rs("COMREPDATE")
        Else
            upd_commit_CD34(6) = Null
        End If
        If sel_rtimelog_rs("COMREPTIME") <> "" Then
            upd_commit_CD34(7) = sel_rtimelog_rs("COMREPTIME")
        Else
            upd_commit_CD34(7) = Null
        End If
        upd_commit_CD34(8) = req_chglen(Trim(hemcbcreport.cmb_reqno.Text))
    End If

    upd_commit_CD34.Execute

    sel_rtimelog_rs.Close
    cmd_commit.Enabled = False
End Sub

Private Sub cmd_print_Click()

CR_CD34.ReportFileName = "C:\DISREPORTSEC\HaematCD34Report.rpt"
CR_CD34.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

CR_CD34.SelectionFormula = "{HAEMATCD34.REQNO} = '" & req_chglen(hemcbcreport.cmb_reqno.Text) & "'"

UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)

CR_CD34.Formulas(0) = "USRID = '" & UI & "'"
CR_CD34.Formulas(1) = "PrePro = '" & frm_CD34.lbl_frmCaption & "'"

CR_CD34.WindowState = crptMaximized
CR_CD34.Action = 1

'**********************************************

End Sub


Private Sub cmd_back_Click()
Unload Me
'DIS_MAIN.show
'hemcbcreport.show
'hemcbcreport.cancelCD34
'hemcbcreport.show
SURFACEFORM.show
End Sub

Private Sub cmd_exit_Click()
End
End Sub


Private Sub Timer1_Timer()

datetime
txt_time.Text = time1

End Sub

