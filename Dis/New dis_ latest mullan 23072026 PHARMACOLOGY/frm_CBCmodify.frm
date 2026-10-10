VERSION 5.00
Begin VB.Form frm_CBCmodify 
   BackColor       =   &H00C0C0FF&
   Caption         =   "CBC Refresh"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form2"
   ScaleHeight     =   9030
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Height          =   3255
      Left            =   1920
      TabIndex        =   0
      Top             =   1920
      Width           =   5655
      Begin VB.TextBox TXT_REQNO 
         Height          =   405
         Left            =   2640
         TabIndex        =   3
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmd_clickit 
         Caption         =   "REFRESH"
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
         Left            =   1200
         TabIndex        =   2
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cmd_back 
         Caption         =   "BACK"
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
         Left            =   3120
         TabIndex        =   1
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisition No."
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
         Left            =   960
         TabIndex        =   4
         Top             =   960
         Width           =   1440
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HAEMATO PATHOLOGY"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   330
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   8595
   End
End
Attribute VB_Name = "frm_CBCmodify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UPDRREQ_QR As rdoQuery
Dim del_dhemcbcdet_qr As rdoQuery
Dim str As String
Dim req As String
Dim ins_hinttfc_qr As rdoQuery
Dim sel_hinter_qr As rdoQuery
Dim sel_hintrfc_qr As rdoQuery
Dim sel_hintrfc_rs As rdoResultset
Dim sel_hinter_rs As rdoResultset

Dim upd_hinter_qr As rdoQuery
Dim upd_hinter_qr1 As rdoQuery


Private Sub cmd_back_Click()
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_clickit_Click()
'On Error GoTo MSG

    If Trim(req) = "" Then
        MsgBox "Please enter requisition number"
        Exit Sub
    End If
    
    If Len(Trim(req)) <> 13 Then
        MsgBox "Please enter proper requisition number"
        Exit Sub
    End If
    
    
    UPDRREQ_QR(0) = req
    UPDRREQ_QR.Execute
    
    del_dhemcbcdet_qr(0) = req
    del_dhemcbcdet_qr.Execute
    
    '***** Insert into log *****
   
    sel_hinter_qr(0) = req
    'sel_hinter_qr.Execute
  
    Set sel_hinter_rs = sel_hinter_qr.OpenResultset(1, 3)
    
    sel_hintrfc_qr(0) = req
    Set sel_hintrfc_rs = sel_hintrfc_qr.OpenResultset(1, 3)
    
    
    
    If Not sel_hinter_rs.EOF Then
        ins_hinttfc_qr(0) = sel_hinter_rs(0)
        ins_hinttfc_qr(1) = sel_hinter_rs(1)
        ins_hinttfc_qr(2) = sel_hinter_rs(2)
        ins_hinttfc_qr(3) = sel_hinter_rs(3)
        ins_hinttfc_qr(4) = sel_hinter_rs(4)
        ins_hinttfc_qr(5) = sel_hinter_rs(5)
        ins_hinttfc_qr(6) = sel_hinter_rs(6)
        ins_hinttfc_qr(7) = sel_hinter_rs(7)
        ins_hinttfc_qr(8) = sel_hinter_rs(8)
        ins_hinttfc_qr(9) = sel_hinter_rs(9)
        ins_hinttfc_qr(10) = sel_hinter_rs(10)
        ins_hinttfc_qr(11) = sel_hinter_rs(11)
        ins_hinttfc_qr(12) = sel_hinter_rs(12)
        ins_hinttfc_qr(13) = sel_hinter_rs(13)
        ins_hinttfc_qr(14) = sel_hinter_rs(14)
        ins_hinttfc_qr(15) = sel_hinter_rs(15)
        ins_hinttfc_qr(16) = sel_hinter_rs(16)
        ins_hinttfc_qr(17) = sel_hinter_rs(17)
        ins_hinttfc_qr(18) = sel_hinter_rs(18)
        ins_hinttfc_qr(19) = sel_hinter_rs(19)
        ins_hinttfc_qr(20) = sel_hinter_rs(20)
        ins_hinttfc_qr(21) = sel_hinter_rs(21)
        ins_hinttfc_qr(22) = sel_hinter_rs(22)
        ins_hinttfc_qr(23) = sel_hinter_rs(23)
        ins_hinttfc_qr(24) = sel_hinter_rs(24)
        ins_hinttfc_qr(25) = sel_hinter_rs(25)
        
        ins_hinttfc_qr(26) = sel_hinter_rs(26)
        ins_hinttfc_qr(27) = sel_hinter_rs(27)
        ins_hinttfc_qr(28) = sel_hinter_rs(28)
        ins_hinttfc_qr(29) = sel_hinter_rs(29)
        ins_hinttfc_qr(30) = sel_hinter_rs(30)
        ins_hinttfc_qr(31) = sel_hinter_rs(31)
        ins_hinttfc_qr(32) = sel_hinter_rs(32)
        ins_hinttfc_qr(33) = sel_hinter_rs(33)
        ins_hinttfc_qr(34) = sel_hinter_rs(34)
        ins_hinttfc_qr(35) = sel_hinter_rs(35)
        ins_hinttfc_qr(36) = sel_hinter_rs(36)
        ins_hinttfc_qr(37) = sel_hinter_rs(37)
        ins_hinttfc_qr(38) = sel_hinter_rs(38)
        ins_hinttfc_qr(39) = sel_hinter_rs(39)
        ins_hinttfc_qr(40) = sel_hinter_rs(40)
        ins_hinttfc_qr(41) = sel_hinter_rs(41)
        ins_hinttfc_qr(42) = sel_hinter_rs(42)
        ins_hinttfc_qr(43) = sel_hinter_rs(43)
        ins_hinttfc_qr(44) = sel_hinter_rs(44)
        ins_hinttfc_qr(45) = sel_hinter_rs(45)
        ins_hinttfc_qr(46) = sel_hinter_rs(46)
        ins_hinttfc_qr(47) = sel_hinter_rs(47)
        ins_hinttfc_qr(48) = sel_hinter_rs(48)
        ins_hinttfc_qr(49) = sel_hinter_rs(49)
        ins_hinttfc_qr(50) = sel_hinter_rs(50)
        
        ins_hinttfc_qr(51) = sel_hinter_rs(51)
        ins_hinttfc_qr(52) = sel_hinter_rs(52)
        ins_hinttfc_qr(53) = sel_hinter_rs(53)
        ins_hinttfc_qr(54) = sel_hinter_rs(54)
        ins_hinttfc_qr(55) = sel_hinter_rs(55)
        ins_hinttfc_qr(56) = sel_hinter_rs(56)
        ins_hinttfc_qr(57) = sel_hinter_rs(57)
        ins_hinttfc_qr(58) = sel_hinter_rs(58)
        ins_hinttfc_qr(59) = sel_hinter_rs(59)
        ins_hinttfc_qr(60) = sel_hinter_rs(60)
        
        'sel_hinter_rs(62) is field 9 'Shraddha 12/01/11
        
        ins_hinttfc_qr(61) = sel_hinter_rs(61)
        ins_hinttfc_qr(62) = sel_hinter_rs(63) 'URGENT
        ins_hinttfc_qr(63) = sel_hinter_rs(64) 'IFLAG
        ins_hinttfc_qr(64) = sel_hinter_rs(65) 'M/C
        ins_hinttfc_qr(65) = sel_hinter_rs(66) 'TTM
        ins_hinttfc_qr(66) = sel_hinter_rs(67) 'TDT
        'ins_hinttfc_qr(67) = sel_hinter_rs(68)  'TDT
        
        'MODDATE  MODTIME  MODBY  IPADD
        
        
        'INSDADDENDUM_QR(2) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        'INSDADDENDUM_QR(3) = time1
        
        ins_hinttfc_qr(67) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
        ins_hinttfc_qr(68) = time1
        ins_hinttfc_qr(69) = UID
        ins_hinttfc_qr(70) = ipadd
        
        
        If sel_hintrfc_rs.EOF Then
            ins_hinttfc_qr.Execute
        End If
        
        upd_hinter_qr(0) = req
        upd_hinter_qr.Execute
            
        upd_hinter_qr1(0) = req
        upd_hinter_qr1.Execute
    
    End If
    '***************************
    MsgBox ("Done......!")
    txt_reqno.Text = ""
    Exit Sub
MSG:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Load()

logonfrm.Main
datetime
str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT='E',COMMITFLAG='' WHERE REQNNO=?"
Set UPDRREQ_QR = rdb.CreateQuery("", str)

str = "DELETE  FROM DISLIB.DHEMCBCDET WHERE REQNO=?"
Set del_dhemcbcdet_qr = rdb.CreateQuery("", str)

'***** Inser into  67
'CASENO  REQNO  LABNO  HB  RBC  HCT  MCV  MCH MCHC CHCM  RDW  HDW  CH  CHDW 14
'PLATELETS  MPV  PDW  PCT  TLC  NEUTROPHIL  LYMPHOCYTE  MONOCYTES  8
'EOSINOPHIL  BASOPHILS  LUC  NRBCS  NEUTROABS  LYMPHOABS  MONOABS  EOSIABS 8
'BASOABS  LUCABS  NRBCSABS  RATIC  RATICABS  MCVR  CHCMR  CHR  R_ANISO  9
'R_MICRO  R_MACRO  R_HCVAR  R_HYPO  R_HYPER  R_RBCF  R _RBCG  W_LS   W_ATYP  9
'W_BLASTS  W_NRBC  W_IG   W_MPO  P_LPLT  P_PLTCLM  FLD1  FLD2   FLD3   FLD4  10
'FLD5   FLD6   FLD7   FLD8    URGENT  IFLAG  MACHINENO  TTIME     TDATE    9
'MODDATE   MODTIME   MODBY                 IPADD
str = "INSERT INTO DISLIB.HINTRFCUP VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
Set ins_hinttfc_qr = rdb.CreateQuery("", str)
'********************************

str = "SELECT * FROM DISLIB.HINTERFACE WHERE REQNO=?"
Set sel_hinter_qr = rdb.CreateQuery("", str)

'***** Shraddha 13/01/11 *****
str = "update interface.HINTERFACE set iflag='1' WHERE REQNO=?"
Set upd_hinter_qr = rdb.CreateQuery("", str)
'*****************************

str = "update dislib.HINTERFACE set iflag='1' WHERE REQNO=?"
Set upd_hinter_qr1 = rdb.CreateQuery("", str)

'DISLIB.HINTRFCUP
str = "SELECT * FROM DISLIB.HINTRFCUP WHERE REQNO=?"
Set sel_hintrfc_qr = rdb.CreateQuery("", str)

'TXT_REQNO.SetFocus
txt_reqno.Text = ""

End Sub

Private Sub txt_reqno_LostFocus()
req = ""
req = UCase(Trim(txt_reqno.Text))
txt_reqno.Text = req
End Sub

