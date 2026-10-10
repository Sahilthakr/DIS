VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form PATHQUERY 
   Caption         =   "Query From Path No."
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12990
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   12990
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
      Begin VB.OptionButton opt_caseno 
         Caption         =   "Case No."
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton opt_path 
         Caption         =   "Path No."
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   7080
      Width           =   11535
      Begin VB.CommandButton cmd_back 
         Caption         =   "&Back"
         Height          =   375
         Left            =   7320
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmb_print 
         Caption         =   "&Print"
         Height          =   375
         Left            =   5760
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmb_view 
         Caption         =   "&View"
         Default         =   -1  'True
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid msf_reqstat 
         Height          =   3255
         Left            =   0
         TabIndex        =   8
         ToolTipText     =   "CLICK ON THE REQUISTION NO."
         Top             =   120
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   8
         BackColor       =   16777215
         ForeColor       =   0
         ForeColorFixed  =   1194390
         WordWrap        =   -1  'True
         Enabled         =   -1  'True
         AllowUserResizing=   1
         FormatString    =   $"PATHQRUERY.frx":0000
      End
   End
   Begin VB.TextBox txt_pathno 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
   Begin MSMask.MaskEdBox msk_caseno 
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      ToolTipText     =   "ENTER THE CASE NUMBER TO DISPLAY PATIENT DETAILS"
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      HideSelection   =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">??/#####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
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
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   11655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Case No."
      Height          =   195
      Left            =   2520
      TabIndex        =   15
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label lbl_name 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5760
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lbl_pathno 
      AutoSize        =   -1  'True
      Caption         =   "Path No."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   765
   End
End
Attribute VB_Name = "PATHQUERY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim seldrequisitn_qr As rdoQuery
Dim seldrequisitn_rs As rdoResultset
Dim SELDREQUISITNFLAG As Boolean

Dim seldrequisitn_qrC As rdoQuery
Dim seldrequisitn_rsC As rdoResultset
Dim SELDREQUISITNCFLAG As Boolean

Dim SELADDENDUMCYQR As rdoQuery
Dim SELADDENDUMCYRS As rdoResultset
Dim SELADDENDUMCYFLAG As Boolean

Dim SELRREQNO_QR As rdoQuery
Dim SELRREQNO_RS As rdoResultset
Dim SELRREQNOFLAG As Boolean
Dim selta101qr As rdoQuery
Dim selta101rs As rdoResultset
Dim SELTA101FLAG As Boolean

Dim SELDFNAC_QR As rdoQuery
Dim SELDFNAC_RS As rdoResultset

Dim selta1012_qr As rdoQuery
Dim selta1012_rs As rdoResultset
Dim SELMA301_QR As rdoQuery
Dim SELMA301_RS As rdoResultset

Dim SELEMPINFO_QR As rdoQuery
Dim SELEMPINFO_RS As rdoResultset
Dim SELEMPINFOFLAG As Boolean

Dim SELDFREPORT_QR As rdoQuery
Dim SELDFREPORT_RS As rdoResultset

Dim SELDADDENDUMQR As rdoQuery
Dim SELDADDENDUMRS As rdoResultset
Dim SELDADDENDUMFLAG As Boolean

Dim SELDREQNOQR As rdoQuery
Dim SELDREQNORS As rdoResultset
Dim SELDREQNOFLAG As Boolean


Dim str As String

Private Sub cmb_print_Click()
'On Error GoTo MSG:
'button_ena_dis ("print")
Dim str1 As String
Dim str2 As String
Dim STR3 As String
 
DoEvents

Dim VAR_REQNNO As String
msf_reqstat.col = 1
str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1
VAR_REQNNO = req_chglen(Left(msf_reqstat.Text, 16))  '''req_chglen(mask_edit_req.Text)

SELRREQNO_QR(0) = VAR_REQNNO
Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
If SELRREQNO_RS.RowCount > 0 Then

End If
SELDREQNOQR(0) = VAR_REQNNO
Set SELDREQNORS = SELDREQNOQR.OpenResultset(1, 3)
If SELRREQNO_RS.RowCount > 0 Then

End If

If SELEMPINFOFLAG = False Then
    SELEMPINFO_QR(0) = msk_caseno.Text
    Set SELEMPINFO_RS = SELEMPINFO_QR.OpenResultset(1, 3)
    SELEMPBASINFOFLAG = True
Else
    SELEMPINFO_QR(0) = msk_caseno.Text
    SELEMPINFO_RS.MoreResults
    SELEMPINFO_RS.Requery
End If

CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"   'AND {RREQUISITN.REQNNO} = '" & VAR_REQNNO & "'"

If Trim(msf_reqstat.TextMatrix(msf_reqstat.row, 7)) = "" Then
    MsgBox "Provisional Report is Not yet Prepared"
    Exit Sub

End If

'*** SURGICAL PAHOLOGY REPORT
''If selrrequisitnrs("modalitycd") = "SP" Then
''
''        If selrrequisitnrs("COMMITFLAG") = "Y" Then
''           SELDADDENDUMQR(0) = VAR_REQNNO
''           Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
''            If SELDADDENDUMRS.RowCount > 0 Then
''                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISSPADD.RPT"
''            Else
''                CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISSP.RPT"
''            End If
''        Else
''            CRYSTALREPORT1.ReportFileName = "C:\DISREPORTSEC\DISSP.RPT"
''        End If
''        CRYSTALREPORT1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
''        seldrequisitn_qr(0) = VAR_REQNNO
''        Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
''        If seldrequisitn_rs.RowCount > 0 Then
''             CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & Trim(seldrequisitn_rs("prepathno")) & "'"
''        Else
''            CRYSTALREPORT1.Formulas(1) = "PREPATH = '" & "" & "'"
''        End If
''        selta1012_qr(0) = selrrequisitnrs("CASENO")
''        Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
''        If selta1012_rs.RowCount > 0 Then
''            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
''        Else
''            CRYSTALREPORT1.Formulas(2) = "PROVDIA = '" & "" & "'"
''        End If
''
''
''        SELMA301_QR(0) = seldrequisitn_rs("sITE")
''        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
''        If SELMA301_RS.RowCount > 0 Then
''            CRYSTALREPORT1.Formulas(3) = "SIT = '" & Trim(SELMA301_RS("SITEDESC")) & "'"
''        Else
''            CRYSTALREPORT1.Formulas(3) = "SIT = '" & "" & "'"
''        End If
''        If seldrequisitn_rs.RowCount > 0 Then
''            CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & Trim(seldrequisitn_rs("OPPROCEDUR")) & "'"
''        Else
''            CRYSTALREPORT1.Formulas(4) = "OPPRO = '" & "" & "'"
''        End If
''
''        If Trim(selrrequisitnrs("REQNSTAT")) = "T" Then
''            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "PROVISIONAL REPORT" & "'"
''        Else
''            CRYSTALREPORT1.Formulas(5) = "PROFIN = '" & "FINAL REPORT" & "'"
''        End If
''        SELDFREPORT_QR(0) = VAR_REQNNO
''        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
''        If SELDFREPORT_RS.RowCount > 0 Then
''            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
''                CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "FROZEN SECTION" & "'"
''            Else
''                CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "" & "'"
''            End If
''        Else
''            CRYSTALREPORT1.Formulas(6) = "FRSEC = '" & "" & "'"
''        End If
''        If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
''            CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
''        Else
''            CRYSTALREPORT1.Formulas(7) = "IHREP = '" & "" & "'"
''        End If
''        CRYSTALREPORT1.Formulas(8) = "USRID = '" & STR & "'"
''        CRYSTALREPORT1.Action = 1
'' End If
 '*** SURGICAL PAHOLOGY REPORT
 '*** IMMUNIOHISTO CHEMISTRY REPORT
 If SELRREQNO_RS("MODALITYCD") = "IH" Then
'    seldrequisitn_rs(0) = VAR_REQNNO
'    Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
    If Len(Trim(SELDREQNORS("SPIHREQNO"))) = 0 Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISIH.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            If SELRREQNO_RS.RowCount > 0 Then
               If Trim(SELRREQNO_RS("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHP REPORT" & "'"
                    'CRYSTALREPORT1.Formulas(3) = "ENT = '" & "ENTERED BY " & cmb_enterby & "'"
                Else
                    CrystalReport1.Formulas(1) = "PROFIN = '" & "HISTOPATH & IHP REPORT" & "'"
                End If
            End If
            selta1012_qr(0) = SELRREQNO_RS("CASENO")
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            If selta1012_rs.RowCount > 0 Then
                CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
            Else
                CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
            End If
            If seldrequisitn_rs.RowCount > 0 Then
                CrystalReport1.Formulas(3) = "PREPATH = '" & Trim(SELDREQNORS("prepathno")) & "'"
            Else
                CrystalReport1.Formulas(3) = "PREPATH = '" & "" & "'"
            End If
            CrystalReport1.Formulas(4) = "USRID = '" & str & "'"
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

            CrystalReport1.Action = 1
      End If
  ' *** IMMUNIO HISTO CHEMISTRY REPORT
  End If
  ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
    If SELRREQNO_RS("MODALITYCD") = "SP" Then
        
        If SELRREQNO_RS("COMMITFLAG") = "Y" Then
            If SELDADDENDUMFLAG = False Then
                SELDADDENDUMQR(0) = VAR_REQNNO
                Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
                SELDADDENDUMFLAG = True
            Else
                SELDADDENDUMQR(0) = VAR_REQNNO
                SELDADDENDUMRS.MoreResults
                SELDADDENDUMRS.Requery
            End If
            If SELDADDENDUMRS.RowCount > 0 Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZENADD.RPT"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
            End If
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\FROZEN.RPT"
        End If
        
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        
        If SELDREQNOFLAG = False Then
            seldrequisitn_qr(0) = VAR_REQNNO
            Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
            SELDREQNOFLAG = True
        Else
            seldrequisitn_qr(0) = VAR_REQNNO
            seldrequisitn_rs.MoreResults
            seldrequisitn_rs.Requery
        End If
            
            If seldrequisitn_rs.RowCount > 0 Then
                CrystalReport1.Formulas(1) = "PREPATH = '" & Trim(SELDREQNORS("prepathno")) & "'"
            Else
                CrystalReport1.Formulas(1) = "PREPATH = '" & "" & "'"
            End If
            selta1012_qr(0) = SELRREQNO_RS("CASENO")
            Set selta1012_rs = selta1012_qr.OpenResultset(1, 3)
            If selta1012_rs.RowCount > 0 Then
                CrystalReport1.Formulas(2) = "PROVDIA = '" & Trim(selta1012_rs("PRVDIAGDES")) & "'"
            Else
                CrystalReport1.Formulas(2) = "PROVDIA = '" & "" & "'"
            End If
        
        
        SELMA301_QR(0) = SELDREQNORS("sITE")
        Set SELMA301_RS = SELMA301_QR.OpenResultset(1, 3)
        If SELMA301_RS.RowCount > 0 Then
            
            CrystalReport1.Formulas(3) = "SIT = '" & Trim(SELMA301_RS("SITEDESC")) & "'"
        Else
            CrystalReport1.Formulas(3) = "SIT = '" & "" & "'"
        End If
        If seldrequisitn_rs.RowCount > 0 Then
            CrystalReport1.Formulas(4) = "OPPRO = '" & Trim(SELDREQNORS("OPPROCEDUR")) & "'"
        Else
            CrystalReport1.Formulas(4) = "OPPRO = '" & "" & "'"
        End If
        str1 = ""
        If Trim(SELRREQNO_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL HISTOPATHOLOGY REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREQNO_RS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            str1 = "Entered By " & Trim(SELRREQNO_RS("enterby"))
            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL HISTOPATHOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        SELDFREPORT_QR(0) = VAR_REQNNO
        Set SELDFREPORT_RS = SELDFREPORT_QR.OpenResultset(1, 3)
        If SELDFREPORT_RS.RowCount > 0 Then
            If Len(Trim(SELDFREPORT_RS("FSPREPORT"))) > 1 Then
                CrystalReport1.Formulas(6) = "FRSEC = '" & "FROZEN SECTION" & "'"
            Else
                CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("IHREPORT"))) > 1 Then
                CrystalReport1.Formulas(7) = "IHREP = '" & "Immunohistochemistry" & "'"
            Else
                CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("grossdesc"))) > 1 Then
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "Grossdescription" & "'"
            Else
                CrystalReport1.Formulas(10) = "GROSDESC = '" & "" & "'"
            End If
            If Len(Trim(SELDFREPORT_RS("spreport"))) > 1 Then
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "Microscopic Description" & "'"
            Else
                CrystalReport1.Formulas(11) = "MICRODESC = '" & "" & "'"
            End If
        Else
            CrystalReport1.Formulas(6) = "FRSEC = '" & "" & "'"
            CrystalReport1.Formulas(7) = "IHREP = '" & "" & "'"
        End If
        
        CrystalReport1.Formulas(8) = "USRID = '" & str & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
 ' *** SURGICAL PATHOLOGY REPORT WITH IMMUNIOHISTOCHEMISTRY REQNO IS INCLUDED
    End If


'*********************** FNAC REPORT*******************************'
    If SELRREQNO_RS("MODALITYCD") = "FA" Then
        If SELRREQNO_RS("COMMITFLAG") = "Y" Then
            If SELDADDENDUMFLAG = False Then
                SELDADDENDUMQR(0) = VAR_REQNNO
                Set SELDADDENDUMRS = SELDADDENDUMQR.OpenResultset(1, 3)
                SELDADDENDUMFLAG = True
            Else
                SELDADDENDUMQR(0) = VAR_REQNNO
                SELDADDENDUMRS.MoreResults
                SELDADDENDUMRS.Requery
            End If
            If SELDADDENDUMRS.RowCount > 0 Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFAADD.RPT"
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFA.RPT"
            End If
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISFA.RPT"
        End If
        
        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        str1 = ""
        
        If Trim(SELRREQNO_RS("REQNSTAT")) = "T" Then
            CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL HISTOPATHOLOGY REPORT" & "'"
            str1 = "Entered By " & Trim(SELRREQNO_RS("enterby"))
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        Else
            str1 = "Entered By " & Trim(SELRREQNO_RS("enterby"))
            CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL HISTOPATHOLOGY REPORT" & "'"
            CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
        End If
        
        SELDFNAC_QR(0) = VAR_REQNNO
        Set SELDFNAC_RS = SELDFNAC_QR.OpenResultset(1, 1)
        
        If SELDFNAC_RS("MICREPORT") = " " Then
            CrystalReport1.Formulas(8) = "fn = '" & "" & "'"
        Else
            CrystalReport1.Formulas(8) = "fn = '" & "12" & "'"
        End If
        
        CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1

    End If
'********************************* FNAC REPORT ENDS************************'


'***CYTOLOGY REPORT****
If SELRREQNO_RS("MODALITYCD") = "CG" Or SELRREQNO_RS("MODALITYCD") = "CN" Then
    
    If opt_path.Value = True Then
        
        If SELRREQNO_RS("COMMITFLAG") = "Y" Then
            
            If SELADDENDUMCYFLAG = False Then
                SELADDENDUMCYQR(0) = VAR_REQNNO
                Set SELADDENDUMCYRS = SELADDENDUMCYQR.OpenResultset(1, 3)
                SELADDENDUMCYFLAG = True
            Else
                SELADDENDUMCYQR(0) = VAR_REQNNO
                SELADDENDUMCYRS.MoreResults
                SELADDENDUMCYRS.Requery
            End If
            
            If SELADDENDUMCYRS.RowCount > 0 Then
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADD.RPT"
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL CYTOLOGY REPORT WITH ADDENDUM" & "'"
                CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
                CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(12) = "NOSLIDES = " & "" & seldrequisitn_rs("LABNO")
            Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCG.RPT"
                If Trim(SELRREQNO_RS("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL CYTOLOGY REPORT" & "'"
                    str1 = ""
                    str1 = "Entered By " & Trim(SELRREQNO_RS("ENTERBY"))
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                Else
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL CYTOLOGY REPORT" & "'"
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                End If
                CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(12) = "NOSLIDES=" & "" & seldrequisitn_rs("LABNO")
            End If
        

            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
        
        Else
                CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCG.RPT"
                If Trim(SELRREQNO_RS("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL CYTOLOGY REPORT" & "'"
                    str1 = ""
                    str1 = "Entered By " & Trim(SELRREQNO_RS("ENTERBY"))
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                Else
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL CYTOLOGY REPORT" & "'"
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                End If
                CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
                CrystalReport1.Formulas(12) = "NOSLIDES=" & "" & seldrequisitn_rs("LABNO")
            
        End If
        
        If Trim(SELRREQNO_RS("PREPORTBY")) <> "" And Trim(SELRREQNO_RS("FREPORTBY")) <> "" And UCase(Trim(SELRREQNO_RS("PREPORTBY"))) <> UCase(Trim(SELRREQNO_RS("FREPORTBY"))) Then
            CrystalReport1.Formulas(11) = "P= '" & Trim(SELRREQNO_RS("PREPORTBY")) & "'"
            CrystalReport1.Formulas(12) = "F= '" & Trim(SELRREQNO_RS("FREPORTBY")) & "'"
            CrystalReport1.Formulas(13) = "PCAP= '" & "Cytotechnologist" & " '"
            CrystalReport1.Formulas(14) = "FCAP= '" & "Pathologist" & " '"
        End If
        If Trim(SELRREQNO_RS("PREPORTBY")) <> "" And Trim(SELRREQNO_RS("FREPORTBY")) = "" Then
            CrystalReport1.Formulas(11) = "PONLY= '" & Trim(SELRREQNO_RS("PREPORTBY")) & "'"
            'CrystalReport1.Formulas(12) = "F= '" & "" & "'"
            CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
        End If
        If Trim(SELRREQNO_RS("PREPORTBY")) <> "" And Trim(SELRREQNO_RS("FREPORTBY")) <> "" And UCase(Trim(SELRREQNO_RS("PREPORTBY"))) = UCase(Trim(SELRREQNO_RS("FREPORTBY"))) Then
            CrystalReport1.Formulas(11) = "PONLY= '" & Trim(SELRREQNO_RS("PREPORTBY")) & "'"
            CrystalReport1.Formulas(13) = "PFCAP= '" & "Cytotechnologist / Pathologist" & " '"
        End If
    Else
        If SELRREQNO_RS("REQNSTAT") <> "F" Then
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGSTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
'            If Trim(TXT_CSITE.Text) = "" Then
'                CrystalReport1.Formulas(1) = "STL = '" & "" & "'"
'            Else
'                CrystalReport1.Formulas(1) = "STL = '" & "Site  " & TXT_CSITE.Text & "'"
'            End If
            If Trim(SELRREQNO_RS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "PROVISIONAL CYTOLOGY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & Trim(SELRREQNO_RS("ENTERBY"))
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "FINAL CYTOLOGY REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
            End If
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        Else
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISCGADDSTAFF.RPT"
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(VAR_REQNNO) & "'"
            CrystalReport1.Formulas(7) = "USRID = '" & str & "'"
        End If
        CrystalReport1.Formulas(8) = "NAM = '" & Trim(SELEMPINFO_RS("TITLE")) & " " & Trim(SELEMPINFO_RS("FNAME")) & " " & Trim(SELEMPINFO_RS("MNAME")) & " " & Trim(SELEMPINFO_RS("LNAME")) & "'"
        CrystalReport1.Formulas(10) = "STDOB = '" & SELEMPINFO_RS("GENDER") & "'"
        CrystalReport1.Formulas(11) = "DOB = " & DateDiff("yyyy", SELEMPINFO_RS("DOB"), SELRREQNO_RS("REQNDT"))
    End If
    CrystalReport1.WindowState = crptMaximized
'    CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"

    CrystalReport1.Action = 1
 End If
 '**** CYTOLOGY REPORT
 
 Exit Sub
'MSG:
'MsgBox Err.Description
End Sub

Private Sub cmb_view_Click()
cmd_cancel.Enabled = True
cmb_print.Enabled = True
cmd_back.Enabled = True
cmb_view.Enabled = False
txt_pathno.Text = ""
txt_name.Text = ""
msf_reqstat.Rows = 1
'txt_pathno.SetFocus
End Sub

Private Sub cmd_back_Click()
cmd_cancel.Enabled = False
cmb_print.Enabled = False
cmd_back.Enabled = True
cmb_view.Enabled = True
txt_pathno.Text = ""
txt_name.Text = ""
msf_reqstat.Rows = 1
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_cancel_Click()
cmd_cancel.Enabled = False
cmb_print.Enabled = False
cmd_back.Enabled = True
cmb_view.Enabled = True
txt_pathno.Text = ""
txt_name.Text = ""
msk_caseno.Text = ""
OPT_CASENO.Value = False
opt_path.Value = False
msf_reqstat.Rows = 1
End Sub

Private Sub Form_Load()
On Error GoTo MSG:
If pathqflag = False Then
SELDREQUISITNFLAG = False
SELRREQNOFLAG = False
SELTA101FLAG = False
pathqflag = True

str = "SELECT * FROM DISLIB.DREQUISITN WHERE PATHNO=?"
Set seldrequisitn_qr = rdb.CreateQuery("seldrequisitn_1qr", str)

str = "SELECT * FROM DISLIB.DREQUISITN WHERE CASENO=?"
Set seldrequisitn_qrC = rdb.CreateQuery("seldrequisitn_1q1r", str)

''STAFF DETAILS
str = "SELECT * FROM EMPLIB.EMPBASINFO WHERE EMPCODE=?"
Set SELEMPINFO_QR = rdb.CreateQuery("SELEMPINFO_QR", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELRREQNO_QR = rdb.CreateQuery("SEL1RREQNO_QR", str)

str = "SELECT * FROM DISLIB.DFNREPORT WHERE REQNO=?"
Set SELDFNAC_QR = rdb.CreateQuery("SELDFNAC_QR", str)

str = "SELECT * FROM PATOOLSLIB.TA101 WHERE CASE_NO=?"
Set selta101qr = rdb.CreateQuery("SEL1TA101QR", str)

str = "select * from patoolslib.ta1012 where case_no=?"
Set selta1012_qr = rdb.CreateQuery("1sel1ta10121_qr1", str)

str = "SELECT * FROM PATOOLSLIB.MA301 WHERE SITECD=?"
Set SELMA301_QR = rdb.CreateQuery("SEL1MA3011_QR1", str)

str = "SELECT * FROM DISLIB.DADDENDUM WHERE REQNO=?"
Set SELDADDENDUMQR = rdb.CreateQuery("SELDADDENDUMQR11", str)

str = "SELECT * FROM DISLIB.ADDENDUMCY WHERE REQNO=?"
Set SELADDENDUMCYQR = rdb.CreateQuery("SELADDENDUMCYQR", str)

str = "SELECT * FROM DISLIB.DREQUISITN WHERE REQNO=?"
Set SELDREQNOQR = rdb.CreateQuery("SELDREQNOQR11", str)

str = "SELECT * FROM DISLIB.DFREPORT WHERE REQNO=?"
Set SELDFREPORT_QR = rdb.CreateQuery("SELDFREPORT11_QR1", str)
cmd_cancel.Enabled = False
cmb_print.Enabled = False
cmd_back.Enabled = True
cmb_view.Enabled = True
txt_pathno.Text = ""
txt_name.Text = ""
msf_reqstat.Rows = 1
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub




Private Sub msk_caseno_LostFocus()
Dim CASENO As String
If OPT_CASENO.Value = True Then
    msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE            |^ PATH NO.    |^                  MEMO                 |^    PAYMENT   |^   COLLECTION    |^       RECEIVED        |^       REPORT"
    If msk_caseno.FormattedText = "" Then
        MsgBox "Case No. is empty"
        Exit Sub
    End If
    CASENO = msk_caseno.Text
    msk_caseno.Text = UCase(padzero(5, "C", CASENO, 3))
        If SELTA101FLAG = False Then
            selta101qr(0) = Trim(msk_caseno.FormattedText)
            Set selta101rs = selta101qr.OpenResultset(1, 3)
            SELTA101FLAG = True
        Else
            selta101qr(0) = Trim(msk_caseno.FormattedText)
            selta101rs.MoreResults
            selta101rs.Requery
        End If
        If selta101rs.RowCount > 0 Then
            txt_name.Text = Trim(selta101rs("tag")) & " " & Trim(selta101rs("first_name")) & "" & Trim(selta101rs("Mid_name")) & " " & Trim(selta101rs("last_name"))
        End If
    
    
    
    If SELDREQUISITNCFLAG = False Then
        seldrequisitn_qrC(0) = UCase(Trim(msk_caseno.FormattedText))
        Set seldrequisitn_rsC = seldrequisitn_qrC.OpenResultset(1, 3)
        SELDREQUISITNCFLAG = True
    Else
        seldrequisitn_qrC(0) = UCase(msk_caseno.FormattedText)
        seldrequisitn_rsC.MoreResults
        seldrequisitn_rsC.Requery
    End If
    If seldrequisitn_rsC.RowCount > 0 Then
       
            i = 1
            While Not seldrequisitn_rsC.EOF
            
                 If SELRREQNOFLAG = False Then
                    SELRREQNO_QR(0) = (seldrequisitn_rsC("reqno"))
                    Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
                    SELRREQNOFLAG = True
                Else
                     SELRREQNO_QR(0) = (seldrequisitn_rsC("reqno"))
                     SELRREQNO_RS.MoreResults
                    SELRREQNO_RS.Requery
                End If
                If SELRREQNO_RS.RowCount > 0 Then
            
            
                    If Trim(SELRREQNO_RS("REQNSTAT")) = "T" Then
                        msf_reqstat.AddItem i & vbTab & SELRREQNO_RS("reqnno") & vbTab & seldrequisitn_rsC("PATHNO") & vbTab & SELRREQNO_RS("memono") & vbTab & SELRREQNO_RS("recptdt") & vbTab & SELRREQNO_RS("APNTDT") & vbTab & SELRREQNO_RS("EXAMINATN") & vbTab & "PRO" & " " & SELRREQNO_RS("PREPORTDT")
                    Else
                        If Trim(SELRREQNO_RS("REQNSTAT")) = "F" Then
                            msf_reqstat.AddItem i & vbTab & SELRREQNO_RS("reqnno") & vbTab & seldrequisitn_rsC("PATHNO") & vbTab & SELRREQNO_RS("memono") & vbTab & SELRREQNO_RS("recptdt") & vbTab & SELRREQNO_RS("APNTDT") & vbTab & SELRREQNO_RS("EXAMINATN") & vbTab & "FIN" & " " & SELRREQNO_RS("PREPORTDT")
                        Else
                            msf_reqstat.AddItem i & vbTab & SELRREQNO_RS("reqnno") & vbTab & seldrequisitn_rsC("PATHNO") & vbTab & SELRREQNO_RS("memono") & vbTab & SELRREQNO_RS("recptdt") & vbTab & SELRREQNO_RS("APNTDT") & vbTab & SELRREQNO_RS("EXAMINATN")
                        End If
                    End If
                    i = i + 1
                    seldrequisitn_rsC.MoveNext
                End If
                'seldrequisitn_rsC.MoveNext
            Wend
        
       End If
End If
End Sub

Private Sub OPT_CASENO_Click()
msk_caseno.SetFocus
End Sub

Private Sub opt_path_Click()
txt_pathno.SetFocus

End Sub

Private Sub txt_pathno_LostFocus()
Dim i As Integer
On Error GoTo MSG:
If opt_path.Value = True Then
    msf_reqstat.FormatString = "SR.NO. |^ REQUISITION NO./ DATE            |^ CASE NO.    |^                  MEMO                 |^    PAYMENT   |^   COLLECTION    |^       RECEIVED        |^       REPORT"
    If Trim(txt_pathno.Text) = "" Then
        MsgBox "Path No. Should Not Be Empty"
    Exit Sub
    End If
    If SELDREQUISITNFLAG = False Then
        seldrequisitn_qr(0) = UCase(Trim(txt_pathno.Text))
        Set seldrequisitn_rs = seldrequisitn_qr.OpenResultset(1, 3)
        SELDREQUISITNFLAG = True
    Else
        seldrequisitn_qr(0) = UCase(txt_pathno.Text)
        seldrequisitn_rs.MoreResults
        seldrequisitn_rs.Requery
    End If
    If seldrequisitn_rs.RowCount > 0 Then
        If SELRREQNOFLAG = False Then
            SELRREQNO_QR(0) = (seldrequisitn_rs("reqno"))
            Set SELRREQNO_RS = SELRREQNO_QR.OpenResultset(1, 3)
            SELRREQNOFLAG = True
        Else
             SELRREQNO_QR(0) = (seldrequisitn_rs("reqno"))
             SELRREQNO_RS.MoreResults
            SELRREQNO_RS.Requery
        End If
        If SELTA101FLAG = False Then
            selta101qr(0) = Trim(seldrequisitn_rs("caseno"))
            Set selta101rs = selta101qr.OpenResultset(1, 3)
            SELTA101FLAG = True
        Else
            selta101qr(0) = Trim(seldrequisitn_rs("caseno"))
            selta101rs.MoreResults
            selta101rs.Requery
        End If
        If selta101rs.RowCount > 0 Then
            txt_name.Text = Trim(selta101rs("tag")) & " " & Trim(selta101rs("first_name")) & "" & Trim(selta101rs("Mid_name")) & " " & Trim(selta101rs("last_name"))
        End If
        If SELRREQNO_RS.RowCount > 0 Then
            i = 1
            While Not SELRREQNO_RS.EOF
                If Trim(SELRREQNO_RS("REQNSTAT")) = "T" Then
                    msf_reqstat.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & SELRREQNO_RS("memono") & vbTab & SELRREQNO_RS("recptdt") & vbTab & SELRREQNO_RS("APNTDT") & vbTab & SELRREQNO_RS("EXAMINATN") & vbTab & "PRO" & " " & SELRREQNO_RS("PREPORTDT")
                Else
                    If Trim(SELRREQNO_RS("REQNSTAT")) = "F" Then
                        msf_reqstat.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & SELRREQNO_RS("memono") & vbTab & SELRREQNO_RS("recptdt") & vbTab & SELRREQNO_RS("APNTDT") & vbTab & SELRREQNO_RS("EXAMINATN") & vbTab & "FIN" & " " & SELRREQNO_RS("PREPORTDT")
                    Else
                        msf_reqstat.AddItem i & vbTab & req_chglen(SELRREQNO_RS("reqnno")) & vbTab & SELRREQNO_RS("caseno") & vbTab & SELRREQNO_RS("memono") & vbTab & SELRREQNO_RS("recptdt") & vbTab & SELRREQNO_RS("APNTDT") & vbTab & SELRREQNO_RS("EXAMINATN")
                    End If
                End If
                i = i + 1
                SELRREQNO_RS.MoveNext
            Wend
        End If
    Else
        MsgBox "Path No. Does not Exists"
    End If
End If
Exit Sub
MSG:
MsgBox Err.Description
End Sub
