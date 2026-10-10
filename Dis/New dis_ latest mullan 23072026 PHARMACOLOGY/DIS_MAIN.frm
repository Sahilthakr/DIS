VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form DIS_MAIN 
   BorderStyle     =   0  'None
   Caption         =   "DIS menu"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   9585
   ForeColor       =   &H00400000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "EmergencyLab Receving"
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   7560
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "MMRU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Microbiology Manual Receiving Unit"
      Top             =   1080
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8760
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   9525
      TabIndex        =   3
      Top             =   7125
      Width           =   9585
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   4320
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label LBL_EXIT 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   4560
      MouseIcon       =   "DIS_MAIN.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6720
      Width           =   510
   End
   Begin VB.Label LBL_CHGPASSWORD 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE &PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   1560
      MouseIcon       =   "DIS_MAIN.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6720
      Width           =   2325
   End
   Begin VB.Label LBL_START 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "&START"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   240
      MouseIcon       =   "DIS_MAIN.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6720
      Width           =   780
   End
   Begin VB.Image Image1 
      Height          =   6375
      Left            =   0
      Picture         =   "DIS_MAIN.frx":091E
      Top             =   0
      Width           =   9510
   End
   Begin VB.Label Label1 
      Caption         =   "DIAGNOSTIC INFORMATION SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Menu MNU_MAS 
      Caption         =   ""
      Begin VB.Menu MNU_MAS1 
         Caption         =   "&MASTER SCREENS"
         Begin VB.Menu INTRA_TRACKING 
            Caption         =   "&INTRA-DEPARTMENTAL TRACKING"
         End
         Begin VB.Menu MNU_ANTI 
            Caption         =   "&ANTIBIOTECTICS MASTER"
         End
         Begin VB.Menu MNU_BIO 
            Caption         =   "&BIOCHEM. TEMPLET MASTER"
         End
         Begin VB.Menu MNU_CCNO 
            Caption         =   "&CC NO & &CC NAME MASTER"
         End
         Begin VB.Menu MNU_INS 
            Caption         =   "&INSTRUCTION MASTER"
            Begin VB.Menu MNU_INS1 
               Caption         =   "&INSTRUCTION MASTER"
            End
            Begin VB.Menu MNU_INSGRP 
               Caption         =   "&INSTRUCTION GROUP MASTER"
            End
         End
         Begin VB.Menu MNUDTM 
            Caption         =   "&DTM SCREEN"
         End
      End
      Begin VB.Menu MNU_LINE 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_req1 
         Caption         =   "&REQUISITION"
         Begin VB.Menu MNU_PATH 
            Caption         =   "R&EQUISITION"
         End
         Begin VB.Menu MNU_APP 
            Caption         =   "APPOINTMENTS"
         End
         Begin VB.Menu MNU_CC 
            Caption         =   "&COLLECTION  CENTRE"
            Begin VB.Menu mnu_old_cc 
               Caption         =   "OLD COLLECTION CENTRE"
            End
            Begin VB.Menu mnuward 
               Caption         =   "&WARD"
            End
            Begin VB.Menu mnuopd 
               Caption         =   "&OPD"
               Begin VB.Menu mnubarcodeprint 
                  Caption         =   "&BARCODE PRINT"
               End
               Begin VB.Menu mnuscanner 
                  Caption         =   "&SCANNER"
               End
            End
         End
         Begin VB.Menu mnu_Daily_Dispatch 
            Caption         =   "DAILY REPORT"
         End
         Begin VB.Menu mnu_CCP 
            Caption         =   "CLINICAL PHARMACOLOGY"
         End
         Begin VB.Menu mnu_dcc 
            Caption         =   "&DEPT.RECEIVING CENTRE"
         End
         Begin VB.Menu mnu_dcc1 
            Caption         =   "&DEPT.RECEIVING CENTRE Barcode"
         End
         Begin VB.Menu mnu_SAMPLERCV 
            Caption         =   "SAMPLE RECEIVE"
            Begin VB.Menu mnu_barCrecv 
               Caption         =   "BARCODE RECEIVE"
            End
            Begin VB.Menu mnu_sample_rcv_list 
               Caption         =   "SAMPLE RCV LIST"
            End
         End
         Begin VB.Menu MNU_EX 
            Caption         =   "&EXAM COMPLETION SCREEN"
         End
      End
      Begin VB.Menu rep 
         Caption         =   "&REPORT"
         Begin VB.Menu BIR 
            Caption         =   "&BIOCHEMICAL INVESTIGATION REPORT"
         End
         Begin VB.Menu mnu_sp 
            Caption         =   "&SURGICAL  PATHOLOGY"
            Begin VB.Menu mnu_report 
               Caption         =   "SURGICAL PATHOLOGY REPORT"
            End
            Begin VB.Menu MNU_PATHBARCODE 
               Caption         =   "PATHOLOGY BARCODE"
            End
            Begin VB.Menu mnu_FZZautocommit 
               Caption         =   "AUTOCOMMIT"
            End
         End
         Begin VB.Menu mnu_imhc 
            Caption         =   "&IMMUNIOHISTOCHEMISTRY  REPORT"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_molep 
            Caption         =   "&MOLECULAR PATHOLOGY"
         End
         Begin VB.Menu mnu_cytogenetics 
            Caption         =   "CYTOGENITICS"
            Begin VB.Menu mnu_cytog 
               Caption         =   "CYTOGENITICS REPORT"
            End
            Begin VB.Menu mnu_VZZautocommit 
               Caption         =   "Autocommit"
            End
         End
         Begin VB.Menu mnu_cytology 
            Caption         =   "CYTOLOGY"
            Begin VB.Menu mnu_CY 
               Caption         =   "&CYTOLOGY REPORT"
            End
            Begin VB.Menu mnu_CYTORPTBARCODE 
               Caption         =   "Cytology Report Barcode"
            End
            Begin VB.Menu mnu_autocommit 
               Caption         =   "AUTOCOMMIT"
            End
         End
         Begin VB.Menu mnu_fnac 
            Caption         =   "&FNAC REPORT"
         End
         Begin VB.Menu MNU_ELAB 
            Caption         =   "EMERGENCY LAB"
            Begin VB.Menu mnu_elabReport 
               Caption         =   "EMERGENCY LAB REPORT"
            End
            Begin VB.Menu MNU_ELIRPT 
               Caption         =   "EMERGENCY LAB INVESTIGATION REPORT"
            End
            Begin VB.Menu MNU_ELAUTOCOMMIT 
               Caption         =   "EMERGENCY LAB AUTOCOMMIT"
            End
            Begin VB.Menu MNU_ELPRINTING 
               Caption         =   "EMERGENCY LAB PRINTING"
            End
         End
         Begin VB.Menu mnu_cp_report 
            Caption         =   "&CLINICAL PHARAMACOLOGY REPORT"
         End
         Begin VB.Menu BC 
            Caption         =   "BIOCHEMISTRY REPORTS"
            Begin VB.Menu MNU_REFRESH 
               Caption         =   "&REFRESH"
               Visible         =   0   'False
            End
            Begin VB.Menu MNU_BC 
               Caption         =   "&BIOCHEMISTRY REPORT"
            End
            Begin VB.Menu MNU_EL 
               Caption         =   "&EMERGENCY LAB REPORT"
               Enabled         =   0   'False
               Visible         =   0   'False
            End
            Begin VB.Menu MNU_EL_autocommit 
               Caption         =   "&EMERGENCY LAB AUTOCOMMIT"
               Enabled         =   0   'False
               Visible         =   0   'False
            End
            Begin VB.Menu MNU_TM1 
               Caption         =   "TUMOUR &MARKER REPORT"
            End
            Begin VB.Menu MNU_BCPRINT 
               Caption         =   "BIOCHEMISTRY PRINTING"
            End
            Begin VB.Menu MNU_ELP 
               Caption         =   "EMERGENCY LAB  PRINTING"
               Enabled         =   0   'False
               Visible         =   0   'False
            End
            Begin VB.Menu MNU_TMP 
               Caption         =   "TUMOUR MARKER PRINTING"
            End
            Begin VB.Menu MNU_MIS 
               Caption         =   "M&IS REPORT"
            End
            Begin VB.Menu MNU_BIO_DISPCH 
               Caption         =   "DISPATCH WARD REPORT"
            End
            Begin VB.Menu MNU_BIO_TEST 
               Caption         =   "Biochemistry Test Count"
            End
         End
         Begin VB.Menu MNU_DTML 
            Caption         =   "&TRANSFUSION MEDICINE REPORTS"
            Begin VB.Menu MNU_DTMLR 
               Caption         =   "&DTM LAB REPORT"
            End
            Begin VB.Menu MNU_BP 
               Caption         =   "&BLOOD PRODUCT REPORT"
            End
            Begin VB.Menu MNU_HLA 
               Caption         =   "&HLA TYPING REPORT"
            End
            Begin VB.Menu mnu_misreport 
               Caption         =   "&MIS REPORT"
            End
            Begin VB.Menu mnu_dtmquery 
               Caption         =   "&DTM QUERY"
            End
            Begin VB.Menu mnu_dtmbatch 
               Caption         =   "&DTM BATCH PRINTING"
            End
         End
         Begin VB.Menu mnu_hp 
            Caption         =   "&HAEMATOPATHOLOGY REPORTS"
            Begin VB.Menu MNU_HeamLblPrint 
               Caption         =   "LabelPrinting"
            End
            Begin VB.Menu MNU_CBCmod 
               Caption         =   "CBC REFRESH"
            End
            Begin VB.Menu MNU_CL 
               Caption         =   "&CLINICAL LABORATPRY"
               Visible         =   0   'False
            End
            Begin VB.Menu MNU_CBC 
               Caption         =   "HAEMATOLOGY  REPORT"
            End
            Begin VB.Menu MNU_BMAP 
               Caption         =   "&BONE MARROW ASPIRATION REPORT"
            End
            Begin VB.Menu MNU_BFC 
               Caption         =   "BODY FLUID CYTOMORPHOLOGY REPORT"
            End
            Begin VB.Menu MNU_HEMATOBATCH 
               Caption         =   "HEMATOLOGY BATCH REPORT"
            End
            Begin VB.Menu MNU_CBCworklist 
               Caption         =   "CBC WORKLIST"
            End
         End
         Begin VB.Menu MNU_ECR 
            Caption         =   "&ECHOCARDIOGRAPHY REPORT"
            Begin VB.Menu mnu_2decho 
               Caption         =   "2D ECHOCARDIOGRAPHY"
            End
            Begin VB.Menu MNU_ECGST 
               Caption         =   "ECG,PFT,STRESS REPORT"
            End
         End
         Begin VB.Menu MNU_ENDO 
            Caption         =   "ENDOSCOPY REPORT"
         End
         Begin VB.Menu MNU_HL 
            Caption         =   "HLA &Typing"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_micro 
            Caption         =   "MICROBIOLOGY REPORT"
            Begin VB.Menu mnuBatchCommit 
               Caption         =   "BATCHCOMMIT"
            End
            Begin VB.Menu MICRO_REPORT1 
               Caption         =   "SEROLOGY REPORT"
            End
            Begin VB.Menu MnuMicroReport 
               Caption         =   "MICROBIOLOGY REPORT"
            End
            Begin VB.Menu MNU_MIC_MIS_REPORT 
               Caption         =   "MIS REPORT"
            End
            Begin VB.Menu MOMICRO_RPT 
               Caption         =   "Molecular Microbiology Report"
            End
            Begin VB.Menu mnu_microbatch 
               Caption         =   "MICROBIOLOGY BATCH REPORT"
            End
            Begin VB.Menu MNU_RAC 
               Caption         =   "SAMPLE ACCESSION RECORD"
            End
            Begin VB.Menu MNU_SAMPLEADD 
               Caption         =   "SAMPLE ADDITION FORM"
            End
            Begin VB.Menu MNU_GZZMANULREC 
               Caption         =   "MANUAL SAMPLE RECEIVING"
            End
            Begin VB.Menu mnu_specUpd 
               Caption         =   "Specimen Update"
            End
         End
         Begin VB.Menu mnu_medonco 
            Caption         =   "MEDICAL ONCOLOGY"
         End
         Begin VB.Menu MNU_TM 
            Caption         =   "&TRANSFUSION MEDICINE REPORT"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu QR 
         Caption         =   "&QUERY"
         Begin VB.Menu REQ_STS 
            Caption         =   "&REQUISITION STATUS"
         End
         Begin VB.Menu mnu_cbcquery 
            Caption         =   "&CBC QUERY"
         End
         Begin VB.Menu MNU_APPQ 
            Caption         =   "&APPOINTMENT QUERY"
         End
         Begin VB.Menu MIS 
            Caption         =   "MIS QUERIES"
            Begin VB.Menu MNU_PATH1 
               Caption         =   "&MIS1"
            End
            Begin VB.Menu MNU_PPATH2 
               Caption         =   "MIS2"
            End
         End
         Begin VB.Menu MNU_PENDING 
            Caption         =   "&PENDING LIST"
            Begin VB.Menu PR 
               Caption         =   "&PATHOLOGY REPORT1"
            End
         End
      End
      Begin VB.Menu mnu_icu_micro 
         Caption         =   "PATHOGEN PROFILE"
      End
      Begin VB.Menu HL7_mnu 
         Caption         =   "&HL7TRANSFER"
      End
      Begin VB.Menu MNU_DIS 
         Caption         =   "&WORKLIST"
      End
      Begin VB.Menu MNU_IOR 
         Caption         =   "&ISOLATED ORG REPORT"
      End
      Begin VB.Menu MNU_PRN 
         Caption         =   "&PRINT SETUP"
      End
   End
End
Attribute VB_Name = "DIS_MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim ScrnMainMenu_Qr As rdoQuery
Dim MAIN_FORM As Boolean
Dim DEPTCODE_QR As rdoQuery
Dim DEPTCODE_RS As rdoResultset

Private Sub BIR_Click()
Me.MousePointer = 11
REPORT_BIOCHEMICAL.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub Command1_Click()
'DMULPROCE.showtm

MICROMANUALREC.show
'frm_PathLblPrint.show
End Sub
Private Sub mnuBatchCommit_Click()
Me.MousePointer = 11
cbccommit.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub Command2_Click()
Barcode.show
End Sub

Private Sub Form_Load()
'On Error GoTo Errhandler
If MAIN_FORM = False Then
    If usr_name = "DISCLK" Then
        MNU_ANTI.Enabled = False
        MNU_BIO.Enabled = False
        MNU_CCNO.Enabled = False
        MNU_INS.Enabled = False
        MNUDTM.Enabled = False
    End If
    If usr_name = "DISADM" Then
        MNU_ANTI.Enabled = True
        MNU_BIO.Enabled = True
        MNU_CCNO.Enabled = True
        MNU_INS.Visible = True
        MNUDTM.Enabled = True
    End If
    If usr_name = "DISCLK" Or usr_name = "DISADM" Then
    
    Else
        str = "SELECT DEPT_CODE FROM DISLIB.DISUSERS WHERE USER_ID=?"
        Set DEPTCODE_QR = rdb.CreateQuery("DEPTCODE_STR", str)
                
        DEPTCODE_QR(0) = LOGIN
        Set DEPTCODE_RS = DEPTCODE_QR.OpenResultset(1, 3)
        USRDEPT = DEPTCODE_RS("DEPT_CODE")
        
        str = "SELECT * FROM DISLIB.DISSCRNS WHERE TRIM(USR_GRP) = ? AND DEPT_CODE=?"
        Set ScrnMainMenu_Qr = rdb.CreateQuery("ScrnMainMenu_Q", str)
        MENU_ENABLE
    End If
    ipadd = Winsock1.LocalIP
    MAIN_FORM = True
End If
Exit Sub
Errhandler:
MsgBox Err.Description, vbCritical, "DIS"
Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LBL_START.ForeColor = vbBlue
LBL_EXIT.ForeColor = vbBlue
LBL_CHGPASSWORD.ForeColor = vbBlue
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub HL7_mnu_Click()
Me.MousePointer = 11
HL7TRANSFER.show
Me.Hide
Me.MousePointer = 0
 
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LBL_START.ForeColor = vbBlack
LBL_EXIT.ForeColor = vbBlack
LBL_CHGPASSWORD.ForeColor = vbBlack
End Sub

Private Sub INTRA_TRACKING_Click()
Me.MousePointer = 11
TRACKING.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_BIODINTER_Click()

Me.MousePointer = 11
BIOCHEM_DINTER.show
Me.Hide
Me.MousePointer = 0



End Sub

Private Sub MNU_BIOSAMREC_Click()
Me.MousePointer = 11
frm_collection_report.show
Me.Hide
Me.MousePointer = 0

End Sub



Private Sub MNU_ECGPDF_Click()
Me.MousePointer = 11
frm_PDFsave.show
Me.Hide
Me.MousePointer = 0

End Sub

Private Sub MNU_ECGPF_Click()
'If Len(Trim(MaskEdBox1.Text)) > 0 Then
    'VAR1 = Trim(Replace(cmb_reqno.Text, "/", ""))    'MaskEdBox1.FormattedText
   ' Dim strLink As String
    'strLink = "http://intranet.tmc.gov.in/emrweb/default.asp" '10.1.4.9
  '  strLink = "http://intranet.tmc.gov.in/EMRREPORTS/allreports.aspx?reqno=" & VAR1 & "~guest"
    'MRUNAL
    strLink = "http://intranet.tmc.gov.in/STRESSTESTREPORT/LOGIN.aspx"
    'END
    Shell "c:\program files\internet explorer\iexplore " & strLink
'
End Sub

Private Sub mnu_sample_rcv_Click()
Me.MousePointer = 11
frm_collection_report.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MOMICRO_RPT_Click()
Me.MousePointer = 11
'frm_collection_report.sh
frm_mo.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub pathmodify_click()
'******** Shraddha Pandit 10/11/08 ******
Me.MousePointer = 11
'frm_collection_report.sh
Form2.show
Me.Hide
Me.MousePointer = 0
'****************************************
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''VARINDER 11012024
Private Sub LBL_CHGPASSWORD_Click()
FRMCHGPWD.show
End Sub

Private Sub LBL_CHGPASSWORD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LBL_CHGPASSWORD.ForeColor = QBColor(2)
End Sub

Private Sub LBL_EXIT_Click()
    End
End Sub

Private Sub LBL_EXIT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LBL_EXIT.ForeColor = QBColor(2)
End Sub

Private Sub LBL_START_Click()
PopupMenu MNU_MAS, , LBL_START.Left - LBL_START.Height, LBL_START.Top - LBL_START.Width
'PopupMenu mnu_req, , 800, 800
End Sub

Private Sub LBL_START_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LBL_START.ForeColor = QBColor(2)
End Sub

Private Sub MICRO_REPORT1_Click()
Me.MousePointer = 11
micro_report.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnu_2decho_Click()
Me.MousePointer = 11
EChocardiologyreport.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_ANTI_Click()
Me.MousePointer = 11
frm_Antibiotic.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_BIO_DISPCH_Click()

Me.MousePointer = 11
FRM_DISPATCH.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnu_cbcquery_Click()
Me.MousePointer = 11
cbc_query.show
Me.Hide
Me.MousePointer = 0
End Sub


Private Sub mnu_CYTORPTBARCODE_Click()
Me.MousePointer = 11
cytology_reportBarcode.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnu_dtmbatch_Click()

Me.MousePointer = 11
frm_dtmbatch.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnu_elabReport_Click()
emg_lab_rpt.show
End Sub

Private Sub MNU_ELAUTOCOMMIT_Click()
If UID = "A400580" Then
    Me.MousePointer = 11
    frm_ELautocommit.show
    Me.Hide
    Me.MousePointer = 0
Else
    MsgBox "Sorry, You are not authorized"
    'MNU_EL_autocommit.Enabled = False
End If
End Sub

Private Sub MNU_ELIRPT_Click()
Me.MousePointer = 11
Report_ELLabInvestign.show
Me.Hide
Me.MousePointer = 0

End Sub

Private Sub MNU_ELPRINTING_Click()
Me.MousePointer = 11
elbatch.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_GZZMANULREC_Click()
Me.MousePointer = 11
MICROMANUALREC.show
Me.Hide
Me.MousePointer = 0

End Sub

Private Sub MNU_HeamLblPrint_Click()
Me.MousePointer = 11
frm_HeamatBarcode.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_ELP_Click()
Me.MousePointer = 11
elbatch.show
Me.Hide
Me.MousePointer = 0
End Sub



Private Sub MNU_PATHBARCODE_Click()
 
Me.MousePointer = 11
frm_PathLblPrint.show
Me.Hide
Me.MousePointer = 0
End Sub


Private Sub MNU_HEMATOBATCH_Click()
Me.MousePointer = 11
frm_hematobatch.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnu_old_cc_Click()
'******** Shraddha Pndit 10/11/08 ******
Me.MousePointer = 11
collection.show
Me.Hide
Me.MousePointer = 0
'****************************************
End Sub


Private Sub mnu_icu_micro_Click()
'******** Shraddha Pandit 10/11/08 ******
Me.MousePointer = 11
frm_ICU_micro.show
Me.Hide
Me.MousePointer = 0
'****************************************
End Sub

Private Sub mnu_Daily_Dispatch_click()
'******** Shraddha Pandit 10/11/08 ******
Me.MousePointer = 11
frm_Daily_rpt.show
'regi.show
Me.Hide
Me.MousePointer = 0
'****************************************
End Sub

Private Sub mnu_sample_rcv_list_click()
'******** Shraddha Pandit 10/11/08 ******
Me.MousePointer = 11
frm_collection_report.show
Me.Hide
Me.MousePointer = 0
'****************************************
End Sub

Private Sub MNU_BIO_TEST_click()
'******** Shraddha Pandit 10/11/08 ******
Me.MousePointer = 11
frm_BiochemTestCount.show
Me.Hide
Me.MousePointer = 0
'****************************************
End Sub
Private Sub mnu_ccp_Click()
Me.MousePointer = 11
frm_ClinicalPharmacologyRCV.show
Me.Hide
Me.MousePointer = 0
End Sub
Private Sub mnu_cp_report_Click()
Me.MousePointer = 11
frm_ClinicalPharmaReport.show
Me.Hide
Me.MousePointer = 0
End Sub


Private Sub MNU_CBCworklist_click()
'******** Shraddha Pandit 13/11/08 ******
Me.MousePointer = 11
frm_CBCworklist.show
Me.Hide
Me.MousePointer = 0
'****************************************
End Sub


Private Sub MNU_SAMPLEADD_click()
'******** Shraddha Pandit 15/04/10 ******
Me.MousePointer = 11
frm_micro_SampleAdd.show
Me.Hide
Me.MousePointer = 0
'****************************************
End Sub

Private Sub MNU_REFRESH_Click()
Me.MousePointer = 11
frm_sero.show
Me.Hide
Me.MousePointer = 0
End Sub

'***** Shraddha 27/10/10 *****
Private Sub MNU_CBCmod_Click()
Me.MousePointer = 11
frm_CBCmodify.show
Me.Hide
Me.MousePointer = 0
End Sub
'*****************************


Private Sub mnu_specUpd_Click()
Me.MousePointer = 11
frm_SpecUpdate.show
Me.Hide
Me.MousePointer = 0
End Sub

'Private Sub MNU_CC_Click()
'Me.MousePointer = 11
'collection.show
'Me.Hide
'Me.MousePointer = 0
'End Sub

Private Sub mnuward_Click()
Me.MousePointer = 11
frmsmpl_coll.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnubarcodeprint_click()
Me.MousePointer = 11
collectionNEW.show
'collectionNEW_1.show
Me.Hide
Me.MousePointer = 0

End Sub
'Private Sub mnuopd_Click()
'Me.MousePointer = 11
'collectionNEW.show
'Me.Hide
'Me.MousePointer = 0
'End Sub
Private Sub mnuscanner_Click()
Me.MousePointer = 11
dept_collection_BARCODE.show
'dept_collection_BARCODE2.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnu_cytog_Click()
Me.MousePointer = 11
cytogenetic_report.show

Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnu_microbatch_Click()
 Me.MousePointer = 11
 frm_microbatch.show
 Me.Hide
 Me.MousePointer = 0
End Sub

'***** Shraddha 15/03/10 *****
Private Sub mnu_medonco_Click()
 Me.MousePointer = 11
 frm_MLreport.show
 Me.Hide
 Me.MousePointer = 0
End Sub
'*****************************

Private Sub mnu_molep_Click()
Me.MousePointer = 11
molecular_report.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_PENDING_Click()
Me.MousePointer = 11
FRM_PENDINGLIST.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_MIC_MIS_REPORT_CLICK()
Me.MousePointer = 11
frm_mis_report.show
Me.Hide
Me.MousePointer = 0
End Sub
'MONO ON 24/04/03
Private Sub MNU_MIS_CLICK()
Me.MousePointer = 11
frm_mis_report.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_RAC_Click()
Me.MousePointer = 11
Smpl_Acc.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnudtm_Click()
Me.MousePointer = 11
frmDTMUserName.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_APP_Click()
Me.MousePointer = 11
tmh_ris_appointment.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_APPQ_Click()
Me.MousePointer = 11
tmh_ris_apptqry.show
Me.Hide
Me.MousePointer = 0

End Sub

Private Sub mnu_bc_Click()
Me.MousePointer = 11
biochemlab_report.show
Me.Hide
Me.MousePointer = 0

End Sub

Private Sub MNU_BCPRINT_Click()
Me.MousePointer = 11
'BATCH_REPORT.show
frm_BCbatchPrint.show   'Commented and added by Shradha on 19/11/10
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_BFC_Click()
Me.MousePointer = 11
'HEMBFREPORT.show
HEMBFREPORTOLD.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_misreport_Click()
Me.MousePointer = 11
frm_mis_report.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_BIO_Click()
Me.MousePointer = 11
dtemplate.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_BMAP_Click()
Me.MousePointer = 11
bmareport.show
Me.Hide
Me.MousePointer = 0

End Sub

Private Sub MNU_BP_Click()
Me.MousePointer = 11
bloodproducts_report.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_CBC_Click()
Me.MousePointer = 11
hemcbcreport.show
Me.Hide
Me.MousePointer = 0
End Sub

'Private Sub mnu_cc_Click()
'Me.MousePointer = 11
'collection.show
'Me.Hide
'Me.MousePointer = 0
'End Sub

Private Sub MNU_CCNO_Click()
Me.MousePointer = 11
Dis_security.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_CL_Click()
Me.MousePointer = 11
clreport.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnu_cy_Click()
Me.MousePointer = 11
'Unload Me
cytology_report.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnu_autocommit_Click()
If UID = "A100035" Or UID = "A400605" Then
    Me.MousePointer = 11
    'Unload Me
    frm_autocommit.show
    Me.Hide
    Me.MousePointer = 0
Else
    MsgBox "You are not authorized to view this report"
End If
End Sub

'***** Shraddha 06/01/10 *****
Private Sub mnu_FZZautocommit_Click()
If UID = "A100035" Then
    Me.MousePointer = 11
    frm_FZZautocommit.show
    Me.Hide
    Me.MousePointer = 0
Else
    MsgBox "You are not authorized to view this report"
End If
End Sub
'*****************************

'***** Shraddha Pandit 17/12/09 *****
Private Sub mnu_VZZautocommit_Click()
Me.MousePointer = 11
'Unload Me
frm_VZZautocommit.show
Me.Hide
Me.MousePointer = 0
End Sub
'************************************
Private Sub mnu_dcc1_Click()
Me.MousePointer = 11
 If Trim(USRDEPT) = "VZZ" Or Trim(USRGRP) = "MEDADM" Then
  dept_collectionnew.show
 Else
   dept_collection_new1.show
   'dept_collection_BARCODE.show
 End If
 
Me.Hide
Me.MousePointer = 0
End Sub
Private Sub mnu_dcc_Click()
Me.MousePointer = 11
 If Trim(USRDEPT) = "VZZ" Or Trim(USRGRP) = "MEDADM" Then
  dept_collectionnew.show
 Else
   dept_collection.show
   'dept_collection_BARCODE.show
 End If
 
Me.Hide
Me.MousePointer = 0
End Sub
Private Sub mnu_barCrecv_Click()
Me.MousePointer = 11
 'dept_collection_BARCODE2.show
 Barcode.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_DIS_Click()
Me.MousePointer = 11
disw.show
Me.Hide
Me.MousePointer = 0
End Sub
'Added on 08/12/08 by Shraddha Pandit
Private Sub MNU_IOR_Click()
If UID = "A100068" Then
    Me.MousePointer = 11
    bact_icu.show
    Me.Hide
    Me.MousePointer = 0
Else
    MsgBox "You are not authorized to view this report"
    
End If
End Sub

Private Sub mnu_dis_inf_Click()
Me.MousePointer = 11
dis_inf.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnu_dtml_Click()
'Unload Me
'dtmlab_report.show
End Sub

Private Sub MNU_ECGTEST_Click()
Me.MousePointer = 11
frm_ecgtest.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_DTMLR_Click()
Me.MousePointer = 11
'DTMREP.show
dtmlab_report.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_DTMRP_Click()
Me.MousePointer = 11
'DTMREP.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_ECGST_Click()
Me.MousePointer = 11
FRM_PFTTEST.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_EL_Click()
emg_lab_rpt.show
End Sub

'***** Shraddha 24/07/10 *****
'MNU_EL_autocommit
Private Sub MNU_EL_autocommit_Click()
'frm_ELautocommit.show

'Added by Shraddha on 05/08/10
If UID = "A400580" Then
    Me.MousePointer = 11
    frm_ELautocommit.show
    Me.Hide
    Me.MousePointer = 0
Else
    MsgBox "Sorry, You are not authorized"
    'MNU_EL_autocommit.Enabled = False
End If
End Sub

'*****************************

Private Sub EL_Click()
Me.MousePointer = 11
elbatch.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_ENDO_Click()
Me.MousePointer = 11
frm_endoscopy1.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_EX_Click()
Me.MousePointer = 11
EXAMCOMP.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnu_fnac_Click()
Me.MousePointer = 11
finnedaspcyt_report.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_HLA_Click()
Me.MousePointer = 11
hlatyping_report.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnu_hp_Click()
'Unload Me
'hematopathology_report.Show
End Sub

Private Sub mnu_imhc_Click()
Me.MousePointer = 11
imhreport.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_INS1_Click()
Me.MousePointer = 11
RISINST.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_INSGRP_Click()
Me.MousePointer = 11
RISLINK.show
Me.Hide
Me.MousePointer = 0
End Sub

'Private Sub mnu_mb_Click()
'Unload Me
'microbiology_report.Show
'Report_microbiology.show
'End Sub

Private Sub mnu_path_Click()
Me.MousePointer = 11

pathology.show
'frm_PathologyRequisitn.show 'Used For Testing Purpose by Shraddha, 08/09/09
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_PATH1_Click()
Me.MousePointer = 11
PATHQUERY.show
Me.Hide
Me.MousePointer = 0
End Sub
Private Sub MNU_dtmquery_Click()
Me.MousePointer = 11
frm_DTM_Query.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_PFTTEST_Click()
Me.MousePointer = 11
FRM_PFTTEST.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_PPATH2_Click()
Me.MousePointer = 11
PATHQUERY.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub MNU_PRN_Click()
CommonDialog1.ShowPrinter

End Sub

Private Sub mnu_report_Click()
Me.MousePointer = 11
SPPATHREPORT.show
Me.Hide
Me.MousePointer = 0
End Sub


Private Sub MNU_STRESSTEST_Click()
Me.MousePointer = 11
FRM_STRESSTEST.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub mnu_tm_Click()
Unload Me
transmed_report.show
End Sub

Private Sub mnumicro_Click()
Report_microbiology.show
End Sub

Private Sub MNU_TM1_Click()
tumour_marker_rpt.show
End Sub

Private Sub MNU_TMP_Click()
Me.MousePointer = 11
tmbatch.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub PR_Click()
Me.MousePointer = 11
Form1.show
Me.Hide
Me.MousePointer = 0
End Sub

Private Sub REQ_STS_Click()
    Me.MousePointer = 11
    tmh_ris_requisition_status1.show
    Me.Hide
    Me.MousePointer = 0
End Sub

Private Sub MnuMicroMaster_click()
    Unload Me
    Load frm_Antibiotic
    frm_Antibiotic.show
End Sub
Private Sub MnuMicroReport_click()
    Me.MousePointer = 11
    Report_microbiology.show
    'MICRO_REPORT.show
    
    Me.Hide
    Me.MousePointer = 0
    
'    Unload Me
'    Load Report_microbiology
'    Report_microbiology.show
End Sub
Public Sub MENU_ENABLE()
Dim RS As rdoResultset

ScrnMainMenu_Qr(0) = Trim(USRGRP)
ScrnMainMenu_Qr(1) = Trim(DEPTCODE_RS("DEPT_CODE"))
Set RS = ScrnMainMenu_Qr.OpenResultset(1, 1)
If RS.RowCount > 0 Then '05/12/08
    RS.MoveFirst '1
    ''SETTING BIOCHEMISTRY BATCH REPORT
        If RS("LEVEL") <> 0 Then
            MNU_BCPRINT.Enabled = True
        Else
            MNU_BCPRINT.Enabled = False
        End If
    RS.MoveNext '2
    ''SETTING BIOCHEMISTRY REPORT
        If RS("LEVEL") <> 0 Then
            MNU_BC.Enabled = True
        Else
            MNU_BC.Enabled = False
        End If
    RS.MoveNext '3
    ''SETTING BLOODPRODUCT REPORT
        If RS("LEVEL") <> 0 Then
            MNU_BP.Enabled = True
        Else
            MNU_BP.Enabled = False
        End If
    RS.MoveNext '4
    ''SETTING BONE MARROW REPORT
        If RS("LEVEL") <> 0 Then
            MNU_BMAP.Enabled = True
        Else
            MNU_BMAP.Enabled = False
        End If
    RS.MoveNext '5
    ''SETTING COLLECTION
        If RS("LEVEL") <> 0 Then
            MNU_CC.Enabled = True
        Else
            MNU_CC.Enabled = False
        End If
    RS.MoveNext '6
    ''SETTING CYTOLOGY REPORT
        If RS("LEVEL") <> 0 Then
            mnu_CY.Enabled = True
        Else
            mnu_CY.Enabled = False
        End If
    RS.MoveNext '7
    ''SETTING DEPARTMENT COLLECTION
        If RS("LEVEL") <> 0 Then
            mnu_dcc.Enabled = True
        Else
            mnu_dcc.Enabled = False
        End If
    RS.MoveNext '8
    ''SETTING DISSECURITY
        If RS("LEVEL") <> 0 Then
            MNU_CCNO.Enabled = True
        Else
            MNU_CCNO.Enabled = False
        End If
    RS.MoveNext '9
    ''SETTING DIS WORK LIST SCREEN
        If RS("LEVEL") <> 0 Then
            MNU_DIS.Enabled = True
        Else
            If Trim(DEPTCODE_RS("DEPT_CODE")) = "UZZ" And Trim(USRGRP) = "SC.OFFICER" Then
            MNU_DIS.Enabled = True
            Else
            MNU_DIS.Enabled = False
            End If
        End If
    RS.MoveNext '10
    ''SETTING BIOCHEMISTRY PROCEDURE MASTER
        If RS("LEVEL") <> 0 Then
            MNU_BIO.Enabled = True
        Else
            MNU_BIO.Enabled = False
        End If
    RS.MoveNext '11
    ''SETTING DTM LAB REPORT
        If RS("LEVEL") <> 0 Then
            MNU_DTMLR.Enabled = True
        Else
            MNU_DTMLR.Enabled = False
        End If
    RS.MoveNext '12
    ''SETTING ECHOCARDIOLOGY REPORT SCREEN
        If RS("LEVEL") <> 0 Then
            mnu_2decho.Enabled = True
        Else
            mnu_2decho.Enabled = False
        End If
    RS.MoveNext '13
    ''SETTING ELBATCH
        If RS("LEVEL") <> 0 Then
            MNU_ELP.Enabled = True
        Else
            MNU_ELP.Enabled = False
        End If
    RS.MoveNext '14
    ''SETTING EMERGENCY LAB REPORT
        If RS("LEVEL") <> 0 Then
            MNU_EL.Enabled = True
        Else
            MNU_EL.Enabled = False
        End If
    RS.MoveNext '15
    ''SETTING EXEME COMPLETION SCREEN
    If RS("LEVEL") <> 0 Then
            MNU_EX.Enabled = True
        Else
            MNU_EX.Enabled = False
        End If
    RS.MoveNext '16
    ''SETTING FNAC REPORT
        If RS("LEVEL") <> 0 Then
            mnu_fnac.Enabled = True
        Else
            mnu_fnac.Enabled = False
        End If
    RS.MoveNext '17
    ''SETTING ANTIMICROBIAL MASTER SCREEN
        If RS("LEVEL") <> 0 Then
            MNU_ANTI.Enabled = True
        Else
            MNU_ANTI.Enabled = False
        End If
    RS.MoveNext '18
    ''SETTING ENDOSCOPY REPORT SCREEN
        If RS("LEVEL") <> 0 Then
            MNU_ENDO.Enabled = True
        Else
            MNU_ENDO.Enabled = False
        End If
    RS.MoveNext '19
    ''SETTING MIS REPORT FOR DTM
        If RS("LEVEL") <> 0 Then
            mnu_misreport.Enabled = True
        Else
            mnu_misreport.Enabled = False
        End If
    RS.MoveNext '20
    ''SETTING MIS REPORT FOR MICROBIOLOGY
        If RS("LEVEL") <> 0 Then
            MNU_MIC_MIS_REPORT.Enabled = True
        Else
            MNU_MIC_MIS_REPORT.Enabled = False
        End If
    RS.MoveNext '21
    ''SETTING MIS REPORT FOR BIOLOGY
        If RS("LEVEL") <> 0 Then
            MNU_MIS.Enabled = True
        Else
            MNU_MIS.Enabled = False
        End If
    RS.MoveNext '22
    ''SETTING PFT TEST REPORT SCREEN
        If RS("LEVEL") <> 0 Then
            MNU_ECGST.Enabled = True
        Else
            MNU_ECGST.Enabled = False
        End If
    RS.MoveNext '23
    ''SETTING DTM USER SCREEN
    If RS("LEVEL") <> 0 Then
            MNUDTM.Enabled = True
        Else
            MNUDTM.Enabled = False
        End If
    RS.MoveNext '24
    ''SETTING BODY FLUID REPORT SCREEN
        If RS("LEVEL") <> 0 Then
            MNU_BFC.Enabled = True
        Else
            MNU_BFC.Enabled = False
        End If
    RS.MoveNext '25
    ''SETTING HAEMATOLOGY REPORT SCREEN
        If RS("LEVEL") <> 0 Or Trim(DEPTCODE_RS("DEPT_CODE")) = "FEC" Then
            MNU_CBC.Enabled = True
        Else
            MNU_CBC.Enabled = False
        End If
    RS.MoveNext '26
     
    ''SETTING HLATYPING REPORT SCREEN
        If RS("LEVEL") <> 0 Then
            MNU_HLA.Enabled = True
        Else
            MNU_HLA.Enabled = False
        End If
    RS.MoveNext '27
    ''SETTING IMMUNOHISTOCHEMISTRY REPORT SCREEN
        If RS("LEVEL") <> 0 Then
            mnu_imhc.Enabled = True
        Else
            mnu_imhc.Enabled = False
        End If
    RS.MoveNext '28
    ''SETTING REQUISITION SCREEN
        If RS("LEVEL") <> 0 Then
            MNU_PATH.Enabled = True
        Else
            MNU_PATH.Enabled = False
        End If
    RS.MoveNext '29
    ''SETTING MIS2 REPORT
        If RS("LEVEL") <> 0 Then
            MNU_PPATH2.Enabled = True
        Else
            MNU_PPATH2.Enabled = False
        End If
    RS.MoveNext '30
    ''SETTING MICROBIOLOBY REPORT SCREEN
        If RS("LEVEL") <> 0 Then
            MnuMicroReport.Enabled = True
        Else
            MnuMicroReport.Enabled = False
        End If
    RS.MoveNext '31
    ''SETTING DIS INSTRUCTION MASTER SCREEN
        If RS("LEVEL") <> 0 Then
            MNU_INS1.Enabled = True
        Else
            MNU_INS1.Enabled = False
        End If
    RS.MoveNext '32
    ''SETTING DIS GROUP INSTRUCTION SCREEN
        If RS("LEVEL") <> 0 Then
            MNU_INSGRP.Enabled = True
        Else
            MNU_INSGRP.Enabled = False
        End If
    RS.MoveNext '33
    ''SETTING SURGICAL PATHOLOGY REPORT SCREEN
        If RS("LEVEL") <> 0 Then
            mnu_report.Enabled = True
        Else
            mnu_report.Enabled = False
        End If
    RS.MoveNext '34
    ''SETTING TUMOUR MARKER BATCH REPORT SCREEN
        If RS("LEVEL") <> 0 Then
            MNU_TMP.Enabled = True
        Else
            MNU_TMP.Enabled = False
        End If
    RS.MoveNext '35
    ''SETTING DIS APPOINTMENT SCREEN
        If RS("LEVEL") <> 0 Then
            MNU_APP.Enabled = True
        Else
            MNU_APP.Enabled = False
        End If
    RS.MoveNext '36
    ''SETTING DIS APPOINTMENT SCREEN MODALITYWISE
        If RS("LEVEL") <> 0 Then
            MNU_APPQ.Enabled = True
        Else
            MNU_APPQ.Enabled = False
        End If
    RS.MoveNext '37
    ''SETTING REQUISITION QUERY SCREEN
        If RS("LEVEL") <> 0 Then
            REQ_STS.Enabled = True
        Else
            REQ_STS.Enabled = False
        End If
    RS.MoveNext '38
    ''SETTING TUMOUR MARKER REPORT SCREEN
        If RS("LEVEL") <> 0 Then
            MNU_TM1.Enabled = True
        Else
            MNU_TM1.Enabled = False
        End If
    RS.MoveNext '39
    ''SETTING PENDING LIST SCREEN
        If RS("LEVEL") <> 0 Then
            MNU_PENDING.Enabled = True
        Else
            MNU_PENDING.Enabled = False
        End If
        mnu_dtmquery.Enabled = False
        ' for disabling biochemistery report when no sub reports are there
        If (MNU_BCPRINT.Enabled = False And MNU_BC.Enabled = False And MNU_EL.Enabled = False And MNU_ELP.Enabled = False And MNU_TM1.Enabled = False And MNU_TMP.Enabled = False) Then
                BC.Enabled = False
        Else
                BC.Enabled = True
        End If
        
Else
    MsgBox "U'r not Authorised to use this Software !!!", vbInformation, "DIS"
    End
    Exit Sub
End If
If Trim(USRDEPT) = "FZZ" Or Trim(USRDEPT) = "ALL" Then

    HL7_mnu.Enabled = True
Else
    HL7_mnu.Enabled = False
End If

    
'If Trim(USRGRP) = "SYSADM" Then
'    mnu_setapass.Enabled = True
'Else
'    mnu_setapass.Enabled = False
'End If
End Sub

