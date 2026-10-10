VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_BCbatchPrint 
   Caption         =   "Biochemistry Batch Print"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form2"
   ScaleHeight     =   4035
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Height          =   1935
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   9135
      Begin VB.CommandButton cmd_back 
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
         Left            =   5880
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txt_fromlabno 
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
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txt_tolabno 
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
         Height          =   285
         Left            =   5640
         TabIndex        =   4
         Top             =   720
         Width           =   1695
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
         Left            =   1920
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
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
         Left            =   3840
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin MSMask.MaskEdBox MSK_TODT1 
         Height          =   375
         Left            =   5640
         TabIndex        =   2
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
      Begin MSMask.MaskEdBox MSK_FROMDT1 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
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
      Begin VB.Label lbl_labno1 
         AutoSize        =   -1  'True
         Caption         =   "From Lab No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   10
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lbl_labno2 
         AutoSize        =   -1  'True
         Caption         =   "To Lab No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   9
         Top             =   720
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
         Left            =   960
         TabIndex        =   8
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
         Left            =   4800
         TabIndex        =   7
         Top             =   360
         Width           =   705
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   480
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label LBL_CAP 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "BIOCHEMISTRY REPORTS"
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
      Left            =   3600
      TabIndex        =   14
      Top             =   600
      Width           =   3735
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
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   4320
      Width           =   5055
   End
   Begin VB.Label LBL_PRINTED 
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   1200
      TabIndex        =   12
      Top             =   5520
      Width           =   3135
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
      Height          =   375
      Left            =   -480
      TabIndex        =   11
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frm_BCbatchPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selma105qr As rdoQuery
Dim selma105rs As rdoResultset
Dim ma105flag As Boolean
Dim str As String
Dim selrrequisitnqr As rdoQuery
Dim selrrequisitnrs As rdoResultset
Dim SELRREQUISITNFLAG As Boolean
Dim SELRREQNOQR As rdoQuery
Dim SELRREQNORS As rdoResultset
Dim SELRREQNOFALG As Boolean
Dim SELRREQNOQRALL As rdoQuery
Dim SELRREQNORSALL As rdoResultset
Dim SELRREQNOFLAGALL As Boolean

Dim selrreqnoipallqr As rdoQuery
Dim selrreqnoipallrs As rdoResultset
Dim selrreqnoipallflag As Boolean

Dim seldmprocedurqr As rdoQuery
Dim seldmprocedurrs As rdoResultset
Dim seldmprocedurflag As Boolean
Dim selpreqnoqr As rdoQuery

'***** Shraddha 19/11/10 *****
Dim selpreqno_IP_qr As rdoQuery
Dim selpreqno_OP_qr As rdoQuery

Dim selpreqno_IP_rs As rdoResultset
Dim selpreqno_OP_rs As rdoResultset
'*****************************

Dim selpreqnors As rdoResultset
Dim selpreqnoflag As Boolean

Dim SELREQNOQR As rdoQuery
Dim SELREQNORS As rdoResultset
Dim SELREQNOFLAG As Boolean

Dim sel_patstat_qr As rdoQuery
Dim sel_patstat_rs As rdoResultset

Dim sel_patstatOPIP_qr As rdoQuery
Dim sel_patstatOPIP_RS As rdoResultset

Dim SEL_STREQNO_QR As rdoQuery
Dim SEL_STREQNO_RS As rdoResultset


Private Sub CMD_PRINT1_Click()
Dim CC As Integer
CC = 0
'On Error GoTo MSG:
If Trim(txt_fromlabno.Text) = "" Then
    MsgBox "Enter the Labno. From"
    Exit Sub
End If
If Trim(txt_tolabno.Text) = "" Then
    MsgBox "Enter the Lab No. upto"
    Exit Sub
End If

'***** Shraddha 26/04/2011 *****

'If Right(txt_fromlabno.Text, 3) >= "601" And Right(txt_tolabno.Text, 3) <= "650" Then
'Commented and added by Shraddha on 16/12/2011
If Right(txt_fromlabno.Text, 4) >= "0601" And Right(txt_tolabno.Text, 4) <= "0650" Then


   

    SEL_STREQNO_QR(0) = Trim(txt_fromlabno.Text)
    SEL_STREQNO_QR(1) = Trim(txt_tolabno.Text)
    Set SEL_STREQNO_RS = SEL_STREQNO_QR.OpenResultset(1, 3)
 SEL_STREQNO_QR.QueryTimeout = 1500
    For i = 0 To 13
        CrystalReport1.Formulas(i) = ""
    Next

    If SEL_STREQNO_RS.RowCount > 0 Then
    
        While Not SEL_STREQNO_RS.EOF
            CC = CC + 1

            var_date_f = Year(MSK_FROMDT1.FormattedText) & "," & month(MSK_FROMDT1.FormattedText) & "," & Day(MSK_FROMDT1.FormattedText)
            var_date_t = Year(MSK_TODT1.FormattedText) & "," & month(MSK_TODT1.FormattedText) & "," & Day(MSK_TODT1.FormattedText)
                 
            CrystalReport1.ReportFileName = "C:\DISREPORTSEC\DISBCSUNI.RPT"
            
            str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
            
            'CrystalReport1.SelectionFormula = "{DISBCUNILF.REQNO}= '" & Trim(SEL_STREQNO_RS("REQNO")) & "' AND ({DISBCUNILF.REQNSTAT}= 'T' OR {DISBCUNILF.REQNSTAT}= 'F')"
            
            CrystalReport1.SelectionFormula = "{RREQUISITN.REQNNO} = '" & Trim(SEL_STREQNO_RS("REQNO")) & "'"
            
            CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(SEL_STREQNO_RS("REQNO")) & "'"
        
            If SELREQNOFLAG = False Then
                SELREQNOQR(0) = SEL_STREQNO_RS("REQNO")
                Set SELREQNORS = SELREQNOQR.OpenResultset(1, 3)
                SELREQNOFLAG = True6
            Else
                SELREQNOQR(0) = SEL_STREQNO_RS("REQNO")
                SELREQNORS.MoreResults
                SELREQNORS.Requery
            End If
            
            If SEL_STREQNO_RS.RowCount > 0 Then
                If Trim(SELREQNORS("REQNSTAT")) = "T" Then
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
                    str1 = ""
                    str1 = "Entered By " & Trim(SELREQNORS("ENTERBY"))
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                 Else
                    CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
                    CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
                 End If
            End If
            CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
            '******
            If seldmprocedurflag = False Then
                seldmprocedurqr(0) = SEL_STREQNO_RS("REQNO")
                Set seldmprocedurrs = seldmprocedurqr.OpenResultset(1, 3)
                seldmprocedurflag = True
            Else
                 seldmprocedurqr(0) = SEL_STREQNO_RS("REQNO")
                 seldmprocedurrs.MoreResults
                 seldmprocedurrs.Requery
            End If
            If seldmprocedurrs.RowCount > 0 Then
                pcount = 0
                While Not seldmprocedurrs.EOF
                    If Trim(seldmprocedurrs("PVALUE")) <> "" Then
                        pcount = pcount + 1
                    End If
                    seldmprocedurrs.MoveNext
                Wend
            End If
            If pcount = 0 Then
                CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
            Else
                CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
            End If
        
            '******
            CrystalReport1.WindowState = crptMaximized
            CrystalReport1.Destination = crptToPrinter
            'CrystalReport1.Destination = crptToWindow
            'CRYSTALREPORT1.WindowState = crptMaximized
            CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
            CrystalReport1.Action = 1
            
            SEL_STREQNO_RS.MoveNext
    Wend
    
End If

If SEL_STREQNO_RS.RowCount = 0 Then
    MsgBox "No records found"
Else
    MsgBox "Printing is completed"
End If
    lbl_cap.Caption = ""
Exit Sub

End If
'********* 26/04/2011 **********




'************** IP ****************
selpreqno_IP_qr.QueryTimeout = 1500

selpreqno_IP_qr(0) = Trim(txt_fromlabno.Text)
selpreqno_IP_qr(1) = Trim(txt_tolabno.Text)
Set selpreqno_IP_rs = selpreqno_IP_qr.OpenResultset(1, 3)

For i = 0 To 13
    CrystalReport1.Formulas(i) = ""
Next

If selpreqno_IP_rs.RowCount > 0 Then
    'Label1.Caption = "No. of Reports in this Combination is " & selpreqno_IP_rs.RowCount
    'LBL_PRINTED.Caption = CC
    While Not selpreqno_IP_rs.EOF
        CC = CC + 1
        'LBL_PRINTED.Caption = CC
'        Dim AA
        
        var_date_f = Year(MSK_FROMDT1.FormattedText) & "," & month(MSK_FROMDT1.FormattedText) & "," & Day(MSK_FROMDT1.FormattedText)
        var_date_t = Year(MSK_TODT1.FormattedText) & "," & month(MSK_TODT1.FormattedText) & "," & Day(MSK_TODT1.FormattedText)
             
        CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISBCUNI.rpt"
        
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
        
        CrystalReport1.SelectionFormula = "{DISBCUNILF.REQNO}= '" & Trim(selpreqno_IP_rs("REQNO")) & "' AND ({DISBCUNILF.REQNSTAT}= 'T' OR {DISBCUNILF.REQNSTAT}= 'F')"

        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selpreqno_IP_rs("REQNO")) & "'"

        If SELREQNOFLAG = False Then
            SELREQNOQR(0) = selpreqno_IP_rs("REQNO")
            Set SELREQNORS = SELREQNOQR.OpenResultset(1, 3)
            SELREQNOFLAG = True6
        Else
            SELREQNOQR(0) = selpreqno_IP_rs("REQNO")
            SELREQNORS.MoreResults
            SELREQNORS.Requery
        End If
        
        If SELREQNORS.RowCount > 0 Then
            If Trim(SELREQNORS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & Trim(SELREQNORS("ENTERBY"))
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             End If
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        '******
        If seldmprocedurflag = False Then
            seldmprocedurqr(0) = selpreqno_IP_rs("REQNO")
            Set seldmprocedurrs = seldmprocedurqr.OpenResultset(1, 3)
            seldmprocedurflag = True
        Else
             seldmprocedurqr(0) = selpreqno_IP_rs("REQNO")
             seldmprocedurrs.MoreResults
             seldmprocedurrs.Requery
        End If
        If seldmprocedurrs.RowCount > 0 Then
            pcount = 0
            While Not seldmprocedurrs.EOF
                If Trim(seldmprocedurrs("PVALUE")) <> "" Then
                    pcount = pcount + 1
                End If
                seldmprocedurrs.MoveNext
            Wend
        End If
        If pcount = 0 Then
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        End If

        '******
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Destination = crptToPrinter
       ' CrystalReport1.Destination = crptToWindow  '02-04-2012
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
     selpreqno_IP_rs.MoveNext
Wend
End If

'************* IP ************

'************* OP ************
selpreqno_OP_qr.QueryTimeout = 50000

selpreqno_OP_qr(0) = Trim(txt_fromlabno.Text)
selpreqno_OP_qr(1) = Trim(txt_tolabno.Text)
Set selpreqno_OP_rs = selpreqno_OP_qr.OpenResultset(1, 3)
    
For i = 0 To 13
    CrystalReport1.Formulas(i) = ""
Next

If selpreqno_OP_rs.RowCount > 0 Then
    'Label1.Caption = "No. of Reports in this Combination is " & selpreqnors.RowCount
    'LBL_PRINTED.Caption = CC
    While Not selpreqno_OP_rs.EOF
        CC = CC + 1
        'LBL_PRINTED.Caption = CC
        'Dim AA
        
        var_date_f = Year(MSK_FROMDT1.FormattedText) & "," & month(MSK_FROMDT1.FormattedText) & "," & Day(MSK_FROMDT1.FormattedText)
        var_date_t = Year(MSK_TODT1.FormattedText) & "," & month(MSK_TODT1.FormattedText) & "," & Day(MSK_TODT1.FormattedText)
        
        CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\DISBCUNI.rpt"
       
        str = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & Format(time1, "hh:MM:SS")
        
        CrystalReport1.SelectionFormula = "{DISBCUNILF.REQNO}= '" & Trim(selpreqno_OP_rs("REQNO")) & "' AND ({DISBCUNILF.REQNSTAT}= 'T' OR {DISBCUNILF.REQNSTAT}= 'F')"

        CrystalReport1.Formulas(0) = "REQNO = '" & req_chglen(selpreqno_OP_rs("REQNO")) & "'"

        If SELREQNOFLAG = False Then
            SELREQNOQR(0) = selpreqno_OP_rs("REQNO")
            Set SELREQNORS = SELREQNOQR.OpenResultset(1, 3)
            SELREQNOFLAG = True6
        Else
            SELREQNOQR(0) = selpreqno_OP_rs("REQNO")
            SELREQNORS.MoreResults
            SELREQNORS.Requery
        End If
        
        If SELREQNORS.RowCount > 0 Then
            If Trim(SELREQNORS("REQNSTAT")) = "T" Then
                CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
                str1 = ""
                str1 = "Entered By " & Trim(SELREQNORS("ENTERBY"))
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             Else
                CrystalReport1.Formulas(5) = "PROFIN = '" & "BIOCHEMISTRY REPORT" & "'"
                CrystalReport1.Formulas(9) = "ENT = '" & str1 & "'"
             End If
        End If
        CrystalReport1.Formulas(12) = "USRID = '" & str & "'"
        '******
        If seldmprocedurflag = False Then
            seldmprocedurqr(0) = selpreqno_OP_rs("REQNO")
            Set seldmprocedurrs = seldmprocedurqr.OpenResultset(1, 3)
            seldmprocedurflag = True
        Else
             seldmprocedurqr(0) = selpreqno_OP_rs("REQNO")
             seldmprocedurrs.MoreResults
             seldmprocedurrs.Requery
        End If
        If seldmprocedurrs.RowCount > 0 Then
            pcount = 0
            While Not seldmprocedurrs.EOF
                If Trim(seldmprocedurrs("PVALUE")) <> "" Then
                    pcount = pcount + 1
                End If
                seldmprocedurrs.MoveNext
            Wend
        End If
        If pcount = 0 Then
            CrystalReport1.Formulas(13) = "PAC = '" & "" & "'"
        Else
            CrystalReport1.Formulas(13) = "PAC = '" & "* Indicate Panic Value" & "'"
        End If

        '******
        CrystalReport1.WindowState = crptMaximized
        CrystalReport1.Destination = crptToPrinter
   '    CrystalReport1.Destination = crptToWindow
        'CRYSTALREPORT1.WindowState = crptMaximized
        CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
        CrystalReport1.Action = 1
     selpreqno_OP_rs.MoveNext
Wend
End If

'*****************************

If selpreqno_IP_rs.RowCount = 0 And selpreqno_OP_rs.RowCount = 0 Then
    MsgBox "No records found"
Else
    MsgBox "Printing is completed"
End If
    lbl_cap.Caption = ""
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_ref_Click()
txt_fromlabno.Text = ""
txt_tolabno.Text = ""
MSK_FROMDT1.Text = ""
MSK_TODT1.Text = ""
End Sub

Private Sub cmd_back_Click()
MSK_FROMDT1.Text = ""
MSK_TODT1.Text = ""
Me.Hide
DIS_MAIN.show
End Sub
Private Sub Form_Load()
Dim str As String

str = "SELECT * FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' AND PATNSTAT=? AND (REQNSTAT= 'T' OR REQNSTAT='F') AND PAT_CATG=? AND PREPORTDT BETWEEN ? AND ?"
Set selrrequisitnqr = rdb.CreateQuery("SELRREQUISITNQRBCR", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE  reqnno=? and PATNSTAT='ST' AND FREPORTDT>=? AND FREPORTDT<=?"
Set sel_patstat_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE  reqnno=? AND (PATNSTAT='IP' OR PATNSTAT='OP') AND FREPORTDT>=? AND FREPORTDT<=?"
Set sel_patstatOPIP_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO IN(SELECT REQNO FROM DISLIB.DREQUISITN WHERE WARD_NAME=?) AND  MODALITYCD='BC' AND PATNSTAT=? AND (REQNSTAT= 'T' OR REQNSTAT='F') AND PAT_CATG=? AND PREPORTDT BETWEEN ? AND ?"
Set SELRREQNOQR = rdb.CreateQuery("SELRREQNOQRBC1 ", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' AND (REQNSTAT= 'T' OR REQNSTAT='F')AND PREPORTDT BETWEEN ? AND ?"
Set SELRREQNOQRALL = rdb.CreateQuery("SELRREQNOQRALLBC", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO IN(SELECT REQNO FROM DISLIB.DREQUISITN)  AND  MODALITYCD='BC' AND PATNSTAT=? AND (REQNSTAT= 'T' OR REQNSTAT='F') AND PAT_CATG=? AND PREPORTDT BETWEEN ? AND ? "  ' order by DISLIB.drequisitn.ward_name "
Set selrreqnoipallqr = rdb.CreateQuery("SELRREQNOQRALLip", str)

str = "SELECT pvalue FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
Set seldmprocedurqr = rdb.CreateQuery("seldmprocedurqrBP", str)


'str = "SELECT REQNO,PATHNO,REFDOC FROM DISLIB.DREQUISITN WHERE REQNO IN (SELECT REQNNO FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' AND (REQNSTAT='T' OR REQNSTAT='F')) AND PATHNO BETWEEN ? AND ?  ORDER BY REFDOC,PATHNO"
str = "SELECT REQNO,PATHNO,REFDOC FROM DISLIB.DREQUISITN WHERE REQNO IN (SELECT REQNNO FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' and PATNSTAT='IP' AND (REQNSTAT='T' OR REQNSTAT='F')) AND PATHNO BETWEEN ? AND ?  ORDER BY PATHNO"
Set selpreqno_IP_qr = rdb.CreateQuery("", str)

'str = "SELECT REQNO,PATHNO,REFTODOC FROM DISLIB.DREQUISITN, RISLIB.RREQUISITN WHERE REQNO IN (SELECT REQNNO FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' and PATNSTAT='OP' AND (REQNSTAT='T' OR REQNSTAT='F')) AND PATHNO BETWEEN ? AND ?  ORDER BY REFBYDOC, PATHNO"
'Commented and added by Shraddha on 04/01/2011
'FOLLOWING QUERY COMMENTED AND NEWLY ADDED BY SHWETA 02-04-2012
'str = "SELECT REQNO,PATHNO,REFTODOC FROM DISLIB.BCBATCHLF WHERE MODALITYCD='BC' and PATHNO BETWEEN ? AND ?  ORDER BY REFBYDOC, PATHNO"
str = "SELECT REQNO,PATHNO,REFTODOC FROM DISLIB.BCBATCHLF WHERE MODALITYCD='BC' AND  PATNSTAT='OP' AND PATHNO BETWEEN ? AND ?  ORDER BY REFBYDOC, PATHNO"
Set selpreqno_OP_qr = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELREQNOQR = rdb.CreateQuery("SELREQNOQRCC", str)

'***** Shraddha 24/04/2011 *****
str = "SELECT REQNO,PATHNO,REFDOC FROM DISLIB.DREQUISITN WHERE REQNO IN (SELECT REQNNO FROM RISLIB.RREQUISITN WHERE MODALITYCD='BC' and PATNSTAT='ST' AND (REQNSTAT='T' OR REQNSTAT='F')) AND PATHNO BETWEEN ? AND ?  ORDER BY PATHNO"
Set SEL_STREQNO_QR = rdb.CreateQuery("", str)
'*******************************

datetime
str = " select ward_name from patoolslib.ma105"
Set selma105qr = rdb.CreateQuery("selma105qrbatrep", str)
End Sub

Private Sub MSK_FROMDT1_GotFocus()
MSK_FROMDT1.Text = Format(DATE11, "dd/mm/yyyy")
End Sub

Private Sub MSK_TODT1_GotFocus()
MSK_TODT1.Text = Format(DATE11, "dd/mm/yyyy")
End Sub

Private Sub txt_fromlabno_LostFocus()
Dim NN As String
If Trim(txt_fromlabno.Text) = "" Then
    Exit Sub
End If
NN = Right(MSK_FROMDT1.FormattedText, 2) & Mid(MSK_FROMDT1.FormattedText, 4, 2) & Left(MSK_FROMDT1.FormattedText, 2)
txt_fromlabno.Text = NN & txt_fromlabno.Text
End Sub

Private Sub txt_tolabno_LostFocus()
Dim NN As String
If Trim(txt_tolabno.Text) = "" Then
    Exit Sub
End If
NN = Right(MSK_TODT1.FormattedText, 2) & Mid(MSK_TODT1.FormattedText, 4, 2) & Left(MSK_TODT1.FormattedText, 2)
txt_tolabno.Text = NN & txt_tolabno.Text
End Sub
