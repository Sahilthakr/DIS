VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_BiochemTestCount 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Biochemical Tests Count"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form2"
   ScaleHeight     =   6915
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_toLabNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txt_frmLabNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      MaxLength       =   3
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ComboBox cmb_mod 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   0
      Top             =   1680
      Width           =   2055
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
      Left            =   1800
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
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
      Left            =   6480
      TabIndex        =   9
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmd_back 
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
      Left            =   4800
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmd_clear 
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
      Left            =   3360
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin MSMask.MaskEdBox msk_frmdt 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin MSMask.MaskEdBox msk_todt 
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin Crystal.CrystalReport CR_CSF 
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CR_SERUM 
      Left            =   600
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CR_URNTEST 
      Left            =   120
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CR_FLUIDS 
      Left            =   600
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "To Lab No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Modality    : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "REPORT FOR BIOCHEMICAL TESTS COUNT"
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
      Left            =   2640
      TabIndex        =   13
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000013&
      Height          =   3135
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   7095
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "To Date : "
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
      Left            =   5520
      TabIndex        =   11
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "From Lab No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
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
      Left            =   720
      TabIndex        =   8
      Top             =   120
      Width           =   8775
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   8280
      Y1              =   3720
      Y2              =   3720
   End
End
Attribute VB_Name = "frm_BiochemTestCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim var_date_f As String
Dim var_date_t As String
Dim dept As String
Dim fdd, fmm, fyy As String
Dim tdd, tmm, tyy As String
Dim i

Dim flabno, tlabno As String

Dim sqlD1 As String
Dim sqlD2 As String
Dim sqlD3 As String
Dim sqlD4 As String

Dim sql2 As String
Dim sql3 As String
Dim sql4 As String
Dim sql5 As String
Dim sql6 As String
Dim sql7 As String
Dim sql8 As String

Dim del_CSF_qr As rdoQuery
Dim del_SERUM_qr As rdoQuery
Dim del_URNTEST_qr As rdoQuery
Dim del_FLUIDS_qr As rdoQuery

Dim sel_rreq_qr As rdoQuery
Dim sel_rreq_rs As rdoResultset

Dim sel_dreq_qr As rdoQuery
Dim sel_dreq_rs As rdoResultset

Dim sel_dmproce_qr As rdoQuery
Dim sel_dmproce_rs As rdoResultset

Dim INS_CSF_QR As rdoQuery
Dim INS_SERUM_QR As rdoQuery
Dim INS_URNTEST_QR As rdoQuery
Dim INS_FLUIDS_QR As rdoQuery
Dim UI As String


Private Sub Form_Load()
cmb_mod.AddItem "Biochemistry$BC"

sqlD1 = "DELETE FROM DISLIB.BIOCSF"
Set del_CSF_qr = rdb.CreateQuery("", sqlD1)

sqlD2 = "DELETE FROM DISLIB.BIOSERUM"
Set del_SERUM_qr = rdb.CreateQuery("", sqlD2)

sqlD3 = "DELETE FROM DISLIB.BIOURNTEST"
Set del_URNTEST_qr = rdb.CreateQuery("", sqlD3)

sqlD4 = "DELETE FROM DISLIB.BIOFLUIDS"
Set del_FLUIDS_qr = rdb.CreateQuery("", sqlD4)

'sql2 = "SELECT * FROM Rislib.Rrequisitn WHERE PREPORTDT >= ? AND PREPORTDT <= ? and MODALITYCD=?"
sql2 = "SELECT * FROM Rislib.Rrequisitn WHERE PREPORTDT >= ? AND PREPORTDT <= ?"
Set sel_rreq_qr = rdb.CreateQuery("", sql2)

sql3 = "SELECT * FROM Dislib.drequisitn WHERE REQNO=? AND PATHNO >= ? AND PATHNO <= ?"
Set sel_dreq_qr = rdb.CreateQuery("", sql3)

sql4 = "SELECT * FROM Dislib.DMPROCEDUR WHERE REQNO=?"
Set sel_dmproce_qr = rdb.CreateQuery("", sql4)

sql5 = "INSERT INTO DISLIB.BIOCSF (CASENO,REQNO,PREPORTDT,GLUCOSE,MTP,LDH,CHLORIDES) VALUES(?,?,?,?,?,?,?)"
Set INS_CSF_QR = rdb.CreateQuery("", sql5)

'sql6 = "INSERT INTO DISLIB.BIOSERUM (CASENO,REQNO,PREPORTDT,HB,CRP,IRON,UIBCTIBC,GGT,CPK,CKMB,LACTATE,MTX,AMICACIN,VANCOMYCIN,CYCLOSPAIN,u,UA,CR,Na,K,Cl,CO2,Ca,PHOS,Mg,ALK,TBIL,DBIL,IBIL,OT,PT,LDH,CHOL,TG,HDL,LDL,AMY,LIP) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
'Commented and added on 26/10/09 by Shraddha Pandit
sql6 = "INSERT INTO DISLIB.BIOSERUM (CASENO,REQNO,PREPORTDT,HB,CRP,IRON,UIBCTIBC,GGT,CPK,CKMB,LACTATE,MTX,AMICACIN,VANCOMYCIN,CYCLOSPAIN,u,UA,CR,Na,K,Cl,CO2,Ca,PHOS,Mg,ALK,TBIL,DBIL,IBIL,OT,PT,LDH,CHOL,TG,HDL,LDL,AMY,LIP,RBS,FBS,PPS) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
'RBS  FBS  PPS

'CASENO reqno PREPORTDT HB CRP IRON UIBCTIBC GGT CPK CKMB LACTATE MTX AMICACIN VANCOMYCIN CYCLOSPAIN 15
'u UA CR Na K Cl CO2 Ca PHOS Mg ALK TBIL DBIL IBIL OT PT LDH CHOL TG HDL LDL AMY LIP 23

Set INS_SERUM_QR = rdb.CreateQuery("", sql6)

sql7 = "INSERT INTO DISLIB.BIOURNTEST (CASENO,REQNO,PREPORTDT,GLUCOSE,UREA,URICACID,CREAT,CALCIUM,PHOSPHO,MTP,NASODIUM,KPOTASSIUM,CLCHLORID) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)"
Set INS_URNTEST_QR = rdb.CreateQuery("", sql7)

sql8 = "INSERT INTO DISLIB.BIOFLUIDS (CASENO,REQNO,PREPORTDT,SODIUM,POTASSIUM,CHLORIDES,MTP,GLUCOSE,UREA,UA,CREATINE,TP,ALB,ALK,TBIL,DBIL,OT,PT,LDH,CA,PHOS,AMYLA,LIPASE,IRON,CHOLES,TRYIGLY,HDL,LDL,GLOBULIN) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
Set INS_FLUIDS_QR = rdb.CreateQuery("", sql8)

datetime

End Sub

Private Sub cmd_print_Click()
If msk_frmdt.Text = "__/__/____" Then
    MsgBox "Please enter From date "
    msk_frmdt.Text = "__/__/____"
    msk_frmdt.SetFocus
    Exit Sub
ElseIf IsDate(msk_frmdt.Text) = False Then
    MsgBox "Please enter valid From date "
    msk_frmdt.Text = "__/__/____"
    msk_frmdt.SetFocus
    Exit Sub
End If

If msk_todt.Text = "__/__/____" Then
    MsgBox "Please enter To date "
    msk_todt.Text = "__/__/____"
    msk_todt.SetFocus
    Exit Sub
ElseIf IsDate(msk_todt.Text) = False Then
    MsgBox "Please enter valid To date "
    msk_todt.Text = "__/__/____"
    msk_todt.SetFocus
    Exit Sub
End If

Dim day_cnt
day_cnt = DateDiff("d", msk_frmdt.Text, msk_todt.Text)
If day_cnt < 0 Then
    MsgBox "PLEASE ENTER VALID DATE RANGE"
    msk_todt.Text = "__/__/____"
    msk_todt.SetFocus
    Exit Sub
End If

'***** Shraddha 20/10/09 *****
If Len(txt_frmLabNo.Text) < 3 Or Len(txt_tolabno.Text) < 3 Then
    MsgBox "Please enter 3 digit Lab No"
    Exit Sub
End If

If CInt(txt_frmLabNo.Text) > CInt(txt_tolabno.Text) Then
    MsgBox "Please enter valid Lab No. range"
    Exit Sub
End If
'*****************************


del_CSF_qr.Execute
del_SERUM_qr.Execute
del_URNTEST_qr.Execute
del_FLUIDS_qr.Execute

var_date_f = Year(msk_frmdt.FormattedText) & "," & month(msk_frmdt.FormattedText) & "," & Day(msk_frmdt.FormattedText)
var_date_t = Year(msk_todt.FormattedText) & "," & month(msk_todt.FormattedText) & "," & Day(msk_todt.FormattedText)

sel_rreq_qr(0) = var_date_f
sel_rreq_qr(1) = var_date_t
i = InStr(cmb_mod, "$")
dept = Mid(cmb_mod.Text, i + 1, 2)

'sel_rreq_qr(2) = dept

Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)

While Not sel_rreq_rs.EOF
    sel_dreq_qr(0) = Trim(sel_rreq_rs("reqnno"))
    
    fdd = Left(msk_frmdt.FormattedText, 2)
    fmm = Mid(msk_frmdt.FormattedText, 4, 2)
    fyy = Right(msk_frmdt.FormattedText, 2)
    
    tdd = Left(msk_todt.FormattedText, 2)
    tmm = Mid(msk_todt.FormattedText, 4, 2)
    tyy = Right(msk_todt.FormattedText, 2)
    
    flabno = fyy & fmm & fdd & Trim(txt_frmLabNo.Text)
    tlabno = tyy & tmm & tdd & Trim(txt_tolabno.Text)
    sel_dreq_qr(1) = Trim(flabno)
    sel_dreq_qr(2) = Trim(tlabno)
    Set sel_dreq_rs = sel_dreq_qr.OpenResultset(1, 3)
    
    If Not sel_dreq_rs.EOF Then
        sel_dmproce_qr(0) = Trim(sel_rreq_rs("reqnno"))
        Set sel_dmproce_rs = sel_dmproce_qr.OpenResultset(1, 3)
        
        '************ CSF *************
        While Not sel_dmproce_rs.EOF
        
        If Trim(sel_dmproce_rs("PRO_CODE")) = "F884" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F885" _
          Or Trim(sel_dmproce_rs("PRO_CODE")) = "F887" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F886" Then
        'CASENO , reqno, PREPORTDT, GLUCOSE, MTP, LDH, CHLORIDES
            INS_CSF_QR(0) = Trim(sel_rreq_rs("CASENO"))
            INS_CSF_QR(1) = Trim(sel_rreq_rs("reqnno"))
            INS_CSF_QR(2) = Trim(sel_rreq_rs("PREPORTDT"))
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F884" Then 'GLUCOSE
                INS_CSF_QR(3) = "YES"
            Else
                INS_CSF_QR(3) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F885" Then 'MTP
                INS_CSF_QR(4) = "YES"
            Else
                INS_CSF_QR(4) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F887" Then 'LDH
                INS_CSF_QR(5) = "YES"
            Else
                INS_CSF_QR(5) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F886" Then 'CHLORIDES
                INS_CSF_QR(6) = "YES"
            Else
                INS_CSF_QR(6) = ""
            End If
            
            INS_CSF_QR.Execute
            
        End If
        '************* End CSF ***************
        
        '************* SERUM *************
        If Trim(sel_dmproce_rs("PRO_CODE")) = "F845" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F829" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F893" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F894" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F999" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F878" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F879" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F892" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F836" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F834" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F835" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F833" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F855" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F856" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F857" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F861" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F862" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F863" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F864" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F888" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F890" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F891" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F869" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F870" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F871" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F872" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F873" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F874" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F876" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F850" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F853" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F851" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F852" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F881" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F882" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F841" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F842" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F843" Then
        'CASENO,REQNO,PREPORTDT,HB,CRP,IRON,UIBCTIBC,GGT,CPK,CKMB,LACTATE,MTX,AMICACIN,VANCOMYCIN,CYCLOSPAIN
        'u UA CR Na K Cl CO2 Ca PHOS Mg ALK TBIL DBIL IBIL OT PT LDH CHOL TG HDL LDL AMY LIP 23
            INS_SERUM_QR(0) = Trim(sel_rreq_rs("CASENO"))
            INS_SERUM_QR(1) = Trim(sel_rreq_rs("reqnno"))
            INS_SERUM_QR(2) = Trim(sel_rreq_rs("PREPORTDT"))
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F845" Then 'HB
                INS_SERUM_QR(3) = "YES"
            Else
                INS_SERUM_QR(3) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F829" Then 'CRP
                INS_SERUM_QR(4) = "YES"
            Else
                INS_SERUM_QR(4) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F893" Then 'IRON
                INS_SERUM_QR(5) = "YES"
            Else
                INS_SERUM_QR(5) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F894" Then 'UIBCTIBC
                INS_SERUM_QR(6) = "YES"
            Else
                INS_SERUM_QR(6) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F999" Then 'GGT
                INS_SERUM_QR(7) = "YES"
            Else
                INS_SERUM_QR(7) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F878" Then 'CPK
                INS_SERUM_QR(8) = "YES"
            Else
                INS_SERUM_QR(8) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F879" Then 'CKMB
                INS_SERUM_QR(9) = "YES"
            Else
                INS_SERUM_QR(9) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F892" Then 'LACTATE
                INS_SERUM_QR(10) = "YES"
            Else
                INS_SERUM_QR(10) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F836" Then 'MTX
                INS_SERUM_QR(11) = "YES"
            Else
                INS_SERUM_QR(11) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F834" Then 'AMICACIN
                INS_SERUM_QR(12) = "YES"
            Else
                INS_SERUM_QR(12) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F835" Then 'VANCOMYCIN
                INS_SERUM_QR(13) = "YES"
            Else
                INS_SERUM_QR(13) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F833" Then 'CYCLOSPAIN
                INS_SERUM_QR(14) = "YES"
            Else
                INS_SERUM_QR(14) = ""
            End If
            
            'u UA CR Na K Cl CO2 Ca PHOS Mg ALK TBIL DBIL IBIL OT PT LDH CHOL TG HDL LDL AMY LIP 23

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F855" Then 'U
                INS_SERUM_QR(15) = "YES"
            Else
                INS_SERUM_QR(15) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F856" Then 'UA
                INS_SERUM_QR(16) = "YES"
            Else
                INS_SERUM_QR(16) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F857" Then 'CR
                INS_SERUM_QR(17) = "YES"
            Else
                INS_SERUM_QR(17) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F861" Then 'Na
                INS_SERUM_QR(18) = "YES"
            Else
                INS_SERUM_QR(18) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F862" Then 'K
                INS_SERUM_QR(19) = "YES"
            Else
                INS_SERUM_QR(19) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F863" Then 'Cl
                INS_SERUM_QR(20) = "YES"
            Else
                INS_SERUM_QR(20) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F864" Then 'CO2
                INS_SERUM_QR(21) = "YES"
            Else
                INS_SERUM_QR(21) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F888" Then 'Ca
                INS_SERUM_QR(22) = "YES"
            Else
                INS_SERUM_QR(22) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F890" Then 'PHOS
                INS_SERUM_QR(23) = "YES"
            Else
                INS_SERUM_QR(23) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F891" Then 'Mg
                INS_SERUM_QR(24) = "YES"
            Else
                INS_SERUM_QR(24) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F869" Then 'Alk
                INS_SERUM_QR(25) = "YES"
            Else
                INS_SERUM_QR(25) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F870" Then 'TBil
                INS_SERUM_QR(26) = "YES"
            Else
                INS_SERUM_QR(26) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F871" Then 'DBil
                INS_SERUM_QR(27) = "YES"
            Else
                INS_SERUM_QR(27) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F872" Then 'IBil
                INS_SERUM_QR(28) = "YES"
            Else
                INS_SERUM_QR(28) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F873" Then 'OT
                INS_SERUM_QR(29) = "YES"
            Else
                INS_SERUM_QR(29) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F874" Then 'PT
                INS_SERUM_QR(30) = "YES"
            Else
                INS_SERUM_QR(30) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F876" Then 'LDH
                INS_SERUM_QR(31) = "YES"
            Else
                INS_SERUM_QR(31) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F850" Then 'Chol
                INS_SERUM_QR(32) = "YES"
            Else
                INS_SERUM_QR(32) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F853" Then 'TG
                INS_SERUM_QR(33) = "YES"
            Else
                INS_SERUM_QR(33) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F851" Then 'HDL
                INS_SERUM_QR(34) = "YES"
            Else
                INS_SERUM_QR(34) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F852" Then 'LDL
                INS_SERUM_QR(35) = "YES"
            Else
                INS_SERUM_QR(35) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F881" Then 'AMY
                INS_SERUM_QR(36) = "YES"
            Else
                INS_SERUM_QR(36) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F882" Then 'LIP
                INS_SERUM_QR(37) = "YES"
            Else
                INS_SERUM_QR(37) = ""
            End If
            
            '***** Shraddha Pandit 26/10/09 RBS  FBS  PPS *****
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F841" Then 'RBS
                INS_SERUM_QR(38) = "YES"
            Else
                INS_SERUM_QR(38) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F842" Then 'FBS
                INS_SERUM_QR(39) = "YES"
            Else
                INS_SERUM_QR(39) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F843" Then 'PPS
                INS_SERUM_QR(40) = "YES"
            Else
                INS_SERUM_QR(40) = ""
            End If
            '************************************
            
            INS_SERUM_QR.Execute

        End If
        '************* End SERUM ***************
        
        '************* URNTEST *************
        If Trim(sel_dmproce_rs("PRO_CODE")) = "F847" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F925" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F926" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F927" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F858" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F931" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F932" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F933" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F928" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F929" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F930" Then
        'CASENO,REQNO,PREPORTDT,GLUCOSE,UREA,URICACID,CREAT,CALCIUM,PHOSPHO,MTP,NASODIUM,KPOTASSIUM,CLCHLORID
            INS_URNTEST_QR(0) = Trim(sel_rreq_rs("CASENO"))
            INS_URNTEST_QR(1) = Trim(sel_rreq_rs("reqnno"))
            INS_URNTEST_QR(2) = Trim(sel_rreq_rs("PREPORTDT"))
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F847" Then 'GLUCOSE
                INS_URNTEST_QR(3) = "YES"
            Else
                INS_URNTEST_QR(3) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F925" Then 'UREA
                INS_URNTEST_QR(4) = "YES"
            Else
                INS_URNTEST_QR(4) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F926" Then 'URICACID
                INS_URNTEST_QR(5) = "YES"
            Else
                INS_URNTEST_QR(5) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F927" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F858" Then 'CREAT
                INS_URNTEST_QR(6) = "YES"
            Else
                INS_URNTEST_QR(6) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F931" Then 'CALCIUM
                INS_URNTEST_QR(7) = "YES"
            Else
                INS_URNTEST_QR(7) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F932" Then 'PHOSPHO
                INS_URNTEST_QR(8) = "YES"
            Else
                INS_URNTEST_QR(8) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F933" Then 'MTP
                INS_URNTEST_QR(9) = "YES"
            Else
                INS_URNTEST_QR(9) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F928" Then 'NASODIUM
                INS_URNTEST_QR(10) = "YES"
            Else
                INS_URNTEST_QR(10) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F929" Then 'KPOTASSIUM
                INS_URNTEST_QR(11) = "YES"
            Else
                INS_URNTEST_QR(11) = ""
            End If

            If Trim(sel_dmproce_rs("PRO_CODE")) = "F930" Then 'CLCHLORID
                INS_URNTEST_QR(12) = "YES"
            Else
                INS_URNTEST_QR(12) = ""
            End If

            INS_URNTEST_QR.Execute

        End If
        '************* End URNTEST ***************
        
        '************* FLUIDS *************
        If Trim(sel_dmproce_rs("PRO_CODE")) = "F937" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F938" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F939" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F885" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F884" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F934" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F935" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F936" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F963" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F964" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F966" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F942" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F941" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F967" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F968" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F973" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F969" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F970" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F971" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F972" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F893" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F943" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F944" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F945" _
        Or Trim(sel_dmproce_rs("PRO_CODE")) = "F946" Or Trim(sel_dmproce_rs("PRO_CODE")) = "F965" Then
        'CASENO,REQNO,PREPORTDT,SODIUM,POTASSIUM,CHLORIDES,MTP,GLUCOSE,UREA,UA,CREATINE,TP,ALB,ALK,TBIL,DBIL,OT,PT,LDH,CA,PHOS,AMYLA,LIPASE,IRON,CHOLES,TRYIGLY,HDL,LDL,GLOBULIN
        INS_FLUIDS_QR(0) = Trim(sel_rreq_rs("CASENO"))
        INS_FLUIDS_QR(1) = Trim(sel_rreq_rs("reqnno"))
        INS_FLUIDS_QR(2) = Trim(sel_rreq_rs("PREPORTDT"))
        
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F937" Then 'SODIUM
                INS_FLUIDS_QR(3) = "YES"
            Else
                INS_FLUIDS_QR(3) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F938" Then 'POTASSIUM
                INS_FLUIDS_QR(4) = "YES"
            Else
                INS_FLUIDS_QR(4) = ""
            End If
                
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F939" Then 'CHLORIDES
                INS_FLUIDS_QR(5) = "YES"
            Else
                INS_FLUIDS_QR(5) = ""
            End If
                
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F885" Then 'MTP
                INS_FLUIDS_QR(6) = "YES"
            Else
                INS_FLUIDS_QR(6) = ""
            End If
                
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F884" Then 'GLUCOSE
                INS_FLUIDS_QR(7) = "YES"
            Else
                INS_FLUIDS_QR(7) = ""
            End If
                
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F934" Then 'UREA
                INS_FLUIDS_QR(8) = "YES"
            Else
                INS_FLUIDS_QR(8) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F935" Then 'UA
                INS_FLUIDS_QR(9) = "YES"
            Else
                INS_FLUIDS_QR(9) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F936" Then 'CREATINE
                INS_FLUIDS_QR(10) = "YES"
            Else
                INS_FLUIDS_QR(10) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F963" Then 'TP
                INS_FLUIDS_QR(11) = "YES"
            Else
                INS_FLUIDS_QR(11) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F964" Then 'ALB
                INS_FLUIDS_QR(12) = "YES"
            Else
                INS_FLUIDS_QR(12) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F966" Then 'ALK
                INS_FLUIDS_QR(13) = "YES"
            Else
                INS_FLUIDS_QR(13) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F942" Then 'TBIL
                INS_FLUIDS_QR(14) = "YES"
            Else
                INS_FLUIDS_QR(14) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F941" Then 'DBIL
                INS_FLUIDS_QR(15) = "YES"
            Else
                INS_FLUIDS_QR(15) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F967" Then 'OT
                INS_FLUIDS_QR(16) = "YES"
            Else
                INS_FLUIDS_QR(16) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F968" Then 'PT
                INS_FLUIDS_QR(17) = "YES"
            Else
                INS_FLUIDS_QR(17) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F973" Then 'LDH
                INS_FLUIDS_QR(18) = "YES"
            Else
                INS_FLUIDS_QR(18) = ""
            End If
                
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F969" Then 'CA
                INS_FLUIDS_QR(19) = "YES"
            Else
                INS_FLUIDS_QR(19) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F970" Then 'PHOS
                INS_FLUIDS_QR(20) = "YES"
            Else
                INS_FLUIDS_QR(20) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F971" Then 'AMYLA
                INS_FLUIDS_QR(21) = "YES"
            Else
                INS_FLUIDS_QR(21) = ""
            End If
                
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F972" Then 'LIPASE
                INS_FLUIDS_QR(22) = "YES"
            Else
                INS_FLUIDS_QR(22) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F893" Then 'IRON
                INS_FLUIDS_QR(23) = "YES"
            Else
                INS_FLUIDS_QR(23) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F943" Then 'CHOLES
                INS_FLUIDS_QR(24) = "YES"
            Else
                INS_FLUIDS_QR(24) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F944" Then 'TRYIGLY
                INS_FLUIDS_QR(25) = "YES"
            Else
                INS_FLUIDS_QR(25) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F945" Then 'HDL
                INS_FLUIDS_QR(26) = "YES"
            Else
                INS_FLUIDS_QR(26) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F946" Then 'LDL
                INS_FLUIDS_QR(27) = "YES"
            Else
                INS_FLUIDS_QR(27) = ""
            End If
            
            If Trim(sel_dmproce_rs("PRO_CODE")) = "F965" Then 'GLOBULIN
                INS_FLUIDS_QR(28) = "YES"
            Else
                INS_FLUIDS_QR(28) = ""
            End If
        
            INS_FLUIDS_QR.Execute
        End If
        
        '************* End FLUIDS ***************
        
        
        sel_dmproce_rs.MoveNext
        Wend
        
    End If
    
    sel_rreq_rs.MoveNext
    
Wend


'*********** Crystal Reports  CSF ****************
CR_CSF.ReportFileName = "C:\DISREPORTSEC\TestCount_BIOCSF.rpt"

UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)

CR_CSF.Formulas(0) = "from=Date(" & var_date_f & ")"
CR_CSF.Formulas(1) = "to=Date(" & var_date_t & ")"

CR_CSF.Formulas(2) = "USRID = '" & UI & "'"

CR_CSF.Formulas(3) = "FromLab = '" & Trim(txt_frmLabNo.Text) & "'"
CR_CSF.Formulas(4) = "ToLab = '" & Trim(txt_tolabno.Text) & "'"
    
CR_CSF.WindowState = crptMaximized
CR_CSF.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CR_CSF.Action = 1

'**************** SERUM *********************
CR_SERUM.ReportFileName = "C:\DISREPORTSEC\TestCount_BIOSERUM.rpt"

UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)

CR_SERUM.Formulas(0) = "from=Date(" & var_date_f & ")"
CR_SERUM.Formulas(1) = "to=Date(" & var_date_t & ")"

CR_SERUM.Formulas(2) = "USRID = '" & UI & "'"

CR_SERUM.Formulas(3) = "FromLab = '" & Trim(txt_frmLabNo.Text) & "'"
CR_SERUM.Formulas(4) = "ToLab = '" & Trim(txt_tolabno.Text) & "'"
    
CR_SERUM.WindowState = crptMaximized
CR_SERUM.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CR_SERUM.Action = 1

'**************** Urine Test *********************
CR_URNTEST.ReportFileName = "C:\DISREPORTSEC\TestCount_BIOURNTEST.rpt"

UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)

CR_URNTEST.Formulas(0) = "from=Date(" & var_date_f & ")"
CR_URNTEST.Formulas(1) = "to=Date(" & var_date_t & ")"

CR_URNTEST.Formulas(2) = "USRID = '" & UI & "'"

CR_URNTEST.Formulas(3) = "FromLab = '" & Trim(txt_frmLabNo.Text) & "'"
CR_URNTEST.Formulas(4) = "ToLab = '" & Trim(txt_tolabno.Text) & "'"
    
CR_URNTEST.WindowState = crptMaximized
CR_URNTEST.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CR_URNTEST.Action = 1

'********************* Fluids *********************
CR_FLUIDS.ReportFileName = "C:\DISREPORTSEC\TestCount_BIOFLUIDS.rpt"

UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)

CR_FLUIDS.Formulas(0) = "from=Date(" & var_date_f & ")"
CR_FLUIDS.Formulas(1) = "to=Date(" & var_date_t & ")"

CR_FLUIDS.Formulas(2) = "USRID = '" & UI & "'"

CR_FLUIDS.Formulas(3) = "FromLab = '" & Trim(txt_frmLabNo.Text) & "'"
CR_FLUIDS.Formulas(4) = "ToLab = '" & Trim(txt_tolabno.Text) & "'"
    
CR_FLUIDS.WindowState = crptMaximized
CR_FLUIDS.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CR_FLUIDS.Action = 1


'********************************************
End Sub

Private Sub cmd_clear_Click()
msk_frmdt.Text = "__/__/____"
msk_todt.Text = "__/__/____"
txt_frmLabNo = ""
txt_tolabno = ""
cmb_mod.Text = ""
msk_frmdt.SetFocus
End Sub

Private Sub cmd_exit_Click()
End
End Sub

Private Sub cmd_back_Click()
DIS_MAIN.show
End Sub


Private Sub txt_frmLabNo_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48) And (KeyAscii <= 57) Or (KeyAscii = 8) Then
    Else
    KeyAscii = 0
    End If
End Sub

Private Sub txt_toLabNo_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48) And (KeyAscii <= 57) Or (KeyAscii = 8) Then
    Else
    KeyAscii = 0
    End If
End Sub

