VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_Daily_rpt 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample Collection Report"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9225
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2880
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
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
      Left            =   4440
      TabIndex        =   6
      Top             =   2760
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
      Left            =   6120
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
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
      Left            =   1080
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.ComboBox cmb_location 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3120
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   1200
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   1200
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
   Begin Crystal.CrystalReport CR 
      Left            =   0
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   7800
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000B&
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
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "Location : "
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
      Left            =   2040
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
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
      Left            =   5160
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000013&
      Height          =   2535
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Collected From Date : "
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
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "frm_Daily_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
' This form has been developed in its entirety by Ms. Shraddha Pandit on 04/05/2009
'*******************************************************************

Dim MODALITY As String
Dim dept As String
Dim var_date_f As String
Dim var_date_t As String

Dim sql As String
Dim sql1 As String
Dim sql2 As String
Dim sql3 As String
Dim sql4 As String
Dim sql5 As String
Dim sql6 As String
Dim sql7 As String

Dim sel_loc_qr As rdoQuery
Dim sel_loc_rs As rdoResultset

Dim sel_rtimelog_qr As rdoQuery
Dim sel_rtimelog_rs As rdoResultset

Dim SEL_PROCODE_QR As rdoQuery
Dim SEL_PROCODE_RS As rdoResultset

Dim SEL_caseno_QR As rdoQuery
Dim SEL_caseno_RS As rdoResultset

Dim SEL_reg_QR As rdoQuery
Dim SEL_reg_RS As rdoResultset

Dim upd_reg_qr As rdoQuery

Dim INS_REGI_QR As rdoQuery

Dim del_regi_qr As rdoQuery

Dim UI As String
Dim cno As String
Dim alp As String
Dim numb As String

Private Sub Form_Load()

'sql1 = "SELECT REQNNO, CASENO, SAMCOLDATE FROM rislib.rtimelog WHERE samcoldate >= ? AND samcoldate <= ?"
'Set sel_rtimelog_qr = rdb.CreateQuery("", sql1)

'sql1 = "SELECT REQNO, CASENO, DATE , lOCATION FROM Dislib.BARCODE WHERE date >= ? AND date <= ? AND LOCATION=?"

'***** 29/03/10 *****
sql1 = "SELECT A.REQNNO, A.CASENO, B.SAMCOLDATE , A.LOCATION FROM RISLIB.RREQUISITN A, rislib.rtimelog B Where a.REQNNO = b.REQNNO AND (B.SAMCOLDATE >= ? AND B.SAMCOLDATE <= ?) AND A.LOCATION=?"
'********************
Set sel_rtimelog_qr = rdb.CreateQuery("", sql1)

sql2 = "DELETE FROM DISLIB.registerop"
Set del_regi_qr = rdb.CreateQuery("", sql2)

sql3 = "INSERT INTO DISLIB.registerop (CASENO, LOCATION, COLLECTDT, BIOCHEM, TUMOUR, CBC, COAG_PROF, ESR, SCREENING, MICROBIOLO) VALUES(?,?,?,?,?,?,?,?,?,?)"
Set INS_REGI_QR = rdb.CreateQuery("", sql3)

sql4 = "SELECT PRO_CODE FROM DISLIB.DMULPROCE WHERE REQNO=?"
Set SEL_PROCODE_QR = rdb.CreateQuery("", sql4)

sql5 = "SELECT caseno FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SEL_caseno_QR = rdb.CreateQuery("", sql5)

sql6 = "SELECT * FROM DISLIB.REGISTEROP WHERE CASENO=?"
Set SEL_reg_QR = rdb.CreateQuery("", sql6)

sql7 = "UPDATE DISLIB.REGISTEROP SET BIOCHEM=?, TUMOUR=?, CBC=?, COAG_PROF=?, ESR=?, SCREENING=?, MICROBIOLO=? WHERE caseno=?"
Set upd_reg_qr = rdb.CreateQuery("", sql7)

sql = "select distinct(location) from dislib.barcode"
Set sel_loc_qr = rdb.CreateQuery("", sql)

datetime

Set sel_loc_rs = sel_loc_qr.OpenResultset(1, 3)
While Not sel_loc_rs.EOF
               
    If Trim(sel_loc_rs(0)) <> "" Then
        cmb_location.AddItem Trim(sel_loc_rs(0))
    End If
    sel_loc_rs.MoveNext
                
Wend
sel_loc_rs.Close

End Sub
Private Sub MaskEdBox1_GotFocus()
MaskEdBox1.Text = Format(Now, "dd/mm/yyyy")
End Sub
Private Sub MaskEdBox2_GotFocus()
MaskEdBox2.Text = Format(Now, "dd/mm/yyyy")
End Sub



Private Sub cmd_back_Click()
DIS_MAIN.show
End Sub

Private Sub cmd_clear_Click()
MaskEdBox1.Text = "__/__/____"
MaskEdBox2.Text = "__/__/____"
MaskEdBox1.SetFocus
End Sub

Private Sub cmd_exit_Click()
End
End Sub

Private Sub cmd_print_Click()

If MaskEdBox1.Text = "__/__/____" Then
        MsgBox "Please enter From date "
        MaskEdBox1.Text = "__/__/____"
        MaskEdBox1.SetFocus
        Exit Sub
ElseIf IsDate(MaskEdBox1.Text) = False Then
        MsgBox "Please enter valid From date "
        MaskEdBox1.Text = "__/__/____"
        MaskEdBox1.SetFocus
        Exit Sub
End If

If MaskEdBox2.Text = "__/__/____" Then
        MsgBox "Please enter To date "
        MaskEdBox2.Text = "__/__/____"
        MaskEdBox2.SetFocus
        Exit Sub
ElseIf IsDate(MaskEdBox2.Text) = False Then
        MsgBox "Please enter valid To date "
        MaskEdBox2.Text = "__/__/____"
        MaskEdBox2.SetFocus
        Exit Sub
End If


    Dim day_cnt
    day_cnt = DateDiff("d", MaskEdBox1.Text, MaskEdBox2.Text)
    If day_cnt < 0 Then
       MsgBox "PLEASE ENTER VALID DATE RANGE"
       MaskEdBox2.Text = "__/__/____"
        MaskEdBox2.SetFocus
        Exit Sub
    End If

    var_date_f = Year(MaskEdBox1.FormattedText) & "," & month(MaskEdBox1.FormattedText) & "," & Day(MaskEdBox1.FormattedText)
    var_date_t = Year(MaskEdBox2.FormattedText) & "," & month(MaskEdBox2.FormattedText) & "," & Day(MaskEdBox2.FormattedText)
   
    '***** 29/03/10 ***** PVT.OPD GEN.OPD
    If Trim(cmb_location.Text) = "GEN.OPD" Then
        'sql1 = "SELECT A.REQNNO, A.CASENO, B.SAMCOLDATE , A.LOCATION FROM RISLIB.RREQUISITN A, rislib.rtimelog B Where a.REQNNO = b.REQNNO AND (B.SAMCOLDATE >= ? AND B.SAMCOLDATE <= ?) AND A.LOCATION=? AND PATNSTAT='OP' AND (PAT_CATG='NC' OR PAT_CATG='C')"
        sql1 = "SELECT A.REQNNO, A.CASENO, B.SAMCOLDATE , A.LOCATION FROM RISLIB.RREQUISITN A, rislib.rtimelog B Where a.REQNNO = b.REQNNO AND (B.SAMCOLDATE >= ? AND B.SAMCOLDATE <= ?) AND PATNSTAT='OP' AND (PAT_CATG='NC' OR PAT_CATG='C')"
    ElseIf Trim(cmb_location.Text) = "PVT.OPD" Then
        'sql1 = "SELECT A.REQNNO, A.CASENO, B.SAMCOLDATE , A.LOCATION FROM RISLIB.RREQUISITN A, rislib.rtimelog B Where a.REQNNO = b.REQNNO AND (B.SAMCOLDATE >= ? AND B.SAMCOLDATE <= ?) AND A.LOCATION=? AND PATNSTAT='OP' AND (PAT_CATG<>'NC' AND PAT_CATG<>'C')"
        sql1 = "SELECT A.REQNNO, A.CASENO, B.SAMCOLDATE , A.LOCATION FROM RISLIB.RREQUISITN A, rislib.rtimelog B Where a.REQNNO = b.REQNNO AND (B.SAMCOLDATE >= ? AND B.SAMCOLDATE <= ?) AND PATNSTAT='OP' AND (PAT_CATG<>'NC' AND PAT_CATG<>'C')"
    'Else
        'sql1 = "SELECT A.REQNNO, A.CASENO, B.SAMCOLDATE , A.LOCATION FROM RISLIB.RREQUISITN A, rislib.rtimelog B Where a.REQNNO = b.REQNNO AND (B.SAMCOLDATE >= ? AND B.SAMCOLDATE <= ?) AND A.LOCATION=?"
    End If
    Set sel_rtimelog_qr = rdb.CreateQuery("", sql1)
    '********************
    sel_rtimelog_qr(0) = var_date_f
    sel_rtimelog_qr(1) = var_date_t
    'sel_rtimelog_qr(2) = cmb_location.Text

    Set sel_rtimelog_rs = sel_rtimelog_qr.OpenResultset(1, 3)

    While Not sel_rtimelog_rs.EOF
    
        dept = Mid(sel_rtimelog_rs("REQNNO"), 1, 3)
        MODALITY = Mid(sel_rtimelog_rs("REQNNO"), 4, 2)
        
        If dept = "FEC" Or dept = "UZZ" Or dept = "NZZ" Or dept = "GZZ" Then
        
        '-------------------------------------------------------------------
            'SEL_reg_QR(0) = sel_rtimelog_rs("REQNNO")
            If Len(Trim(sel_rtimelog_rs("CASENO"))) <> 6 Then
                alp = Left(Trim(sel_rtimelog_rs("CASENO")), 2)
                numb = Right(Trim(sel_rtimelog_rs("CASENO")), 5)
                cno = alp & "/" & numb
                SEL_reg_QR(0) = cno
            Else
                SEL_reg_QR(0) = Trim(sel_rtimelog_rs("CASENO"))
            End If
            
            'SEL_reg_QR(0) = sel_rtimelog_rs("CASENO")
            Set SEL_reg_RS = SEL_reg_QR.OpenResultset(1, 3)
            
            If SEL_reg_RS.EOF Or SEL_reg_RS.BOF Then
        '-------------------------------------------------------------------
                'If Trim(sel_rtimelog_rs("CASENO")) = "" Then
                '    SEL_caseno_QR(0) = sel_rtimelog_rs("REQNNO")
                '    Set SEL_caseno_RS = SEL_caseno_QR.OpenResultset(1, 3)
                '    INS_REGI_QR(0) = Trim(SEL_caseno_RS("CASENO"))
                'Else
                If Len(Trim(sel_rtimelog_rs("CASENO"))) <> 6 Then
                    'alp = Left(Trim(sel_rtimelog_rs("CASENO")), 2)
                    'numb = Right(Trim(sel_rtimelog_rs("CASENO")), 5)
                    'cno = alp & "/" & numb
                    'INS_REGI_QR(0) = Trim(sel_rtimelog_rs("CASENO"))
                    INS_REGI_QR(0) = cno
                Else
                    INS_REGI_QR(0) = Trim(sel_rtimelog_rs("CASENO"))
                End If
               ' End If
                
                'INS_REGI_QR(1) = Trim(sel_rtimelog_rs("REQNNO"))
                'INS_REGI_QR(1) = ""
                'INS_REGI_QR(1) = Trim(sel_rtimelog_rs("LOCATION"))
                If IsNull(Trim(sel_rtimelog_rs("LOCATION"))) = False And Trim(sel_rtimelog_rs("LOCATION")) <> "" Then
                    INS_REGI_QR(1) = Trim(sel_rtimelog_rs("LOCATION"))
                Else
                    INS_REGI_QR(1) = ""
                End If
                INS_REGI_QR(2) = sel_rtimelog_rs("SAMCOLDATE")
                
                If MODALITY = "BC" Then
                    INS_REGI_QR(3) = "YES"
                Else
                    INS_REGI_QR(3) = ""
                End If
                
                If MODALITY = "TM" Then
                    INS_REGI_QR(4) = "YES"
                Else
                    INS_REGI_QR(4) = ""
                End If
        
                'TO SELECT PRO_CODE
                SEL_PROCODE_QR(0) = sel_rtimelog_rs("REQNNO")
                Set SEL_PROCODE_RS = SEL_PROCODE_QR.OpenResultset(1, 3)
        
                While Not SEL_PROCODE_RS.EOF
        
                    If SEL_PROCODE_RS(0) = "U721" Then
                        INS_REGI_QR(5) = "YES" 'CBC
                    Else
                        INS_REGI_QR(5) = ""
                    End If
        
                    If SEL_PROCODE_RS(0) = "U709" Then
                        INS_REGI_QR(6) = "YES" 'Coag_prof
                    Else
                        INS_REGI_QR(6) = ""
                    End If
        
                    If SEL_PROCODE_RS(0) = "U706" Then
                        INS_REGI_QR(7) = "YES" 'ESR
                    Else
                        INS_REGI_QR(7) = ""
                    End If
        
                    SEL_PROCODE_RS.MoveNext
                Wend
        
                SEL_PROCODE_RS.Close
        
                If dept = "NZZ" And MODALITY = "SS" Then
                    INS_REGI_QR(8) = "YES" 'screening
                Else
                    INS_REGI_QR(8) = ""
                End If
                
                If dept = "GZZ" And MODALITY = "SE" Then
                    INS_REGI_QR(9) = "YES" 'microbio
                Else
                    INS_REGI_QR(9) = ""
                End If
        
                INS_REGI_QR.Execute
            '------------------------------------------------------------
            Else
            'upd_reg_qr
                If MODALITY = "BC" Then
                    upd_reg_qr(0) = "YES"
                Else
                    upd_reg_qr(0) = SEL_reg_RS("BIOCHEM")
                End If
                
                If MODALITY = "TM" Then
                    upd_reg_qr(1) = "YES"
                Else
                    upd_reg_qr(1) = SEL_reg_RS("TUMOUR")
                End If
        
                'TO SELECT PRO_CODE
                SEL_PROCODE_QR(0) = sel_rtimelog_rs("REQNNO")
                Set SEL_PROCODE_RS = SEL_PROCODE_QR.OpenResultset(1, 3)
        
                While Not SEL_PROCODE_RS.EOF
        
                    If SEL_PROCODE_RS(0) = "U721" Then
                        upd_reg_qr(2) = "YES" 'CBC
                    Else
                        upd_reg_qr(2) = SEL_reg_RS("CBC")
                    End If
        
                    If SEL_PROCODE_RS(0) = "U709" Then
                        upd_reg_qr(3) = "YES" 'Coag_prof
                    Else
                        upd_reg_qr(3) = SEL_reg_RS("COAG_PROF")
                    End If
        
                    If SEL_PROCODE_RS(0) = "U706" Then
                        upd_reg_qr(4) = "YES" 'ESR
                    Else
                        upd_reg_qr(4) = SEL_reg_RS("ESR")
                    End If
        
                    SEL_PROCODE_RS.MoveNext
                Wend
        
                SEL_PROCODE_RS.Close
        
                If dept = "NZZ" And MODALITY = "SS" Then
                    upd_reg_qr(5) = "YES" 'screening
                Else
                    upd_reg_qr(5) = SEL_reg_RS("SCREENING")
                End If
                
                If dept = "GZZ" And MODALITY = "SE" Then
                    upd_reg_qr(6) = "YES" 'microbio
                Else
                    upd_reg_qr(6) = SEL_reg_RS("MICROBIOLO")
                End If
                
                
                If Len(Trim(sel_rtimelog_rs("CASENO"))) <> 6 Then
                    upd_reg_qr(7) = cno
                Else
                    upd_reg_qr(7) = Trim(sel_rtimelog_rs("CASENO"))
                End If
                
                'upd_reg_qr(7) = sel_rtimelog_rs("CASENO")
                
                upd_reg_qr.Execute
            End If
            '------------------------------------------------------------
        End If
        
        sel_rtimelog_rs.MoveNext
        
    Wend

    sel_rtimelog_rs.Close

    CR.ReportFileName = "C:\DISREPORTSEC\SampleDaily_Dispatch.rpt"
    UI = ipadd & " " & Format(DATE11, "DD/MM/YYYY") & " " & time1 & " " & Right(LOGIN, 6)
    CR.Formulas(0) = "from=Date(" & var_date_f & ")"
    CR.Formulas(1) = "to=Date(" & var_date_t & ")"
    
    CR.Formulas(2) = "USRID = '" & UI & "'"
    
    'CR.SelectionFormula = "{DMULPROCE.REQNDT}>=Date(" & var_date_f & ") AND {DMULPROCE.REQNDT}<=Date(" & var_date_t & ")"
    CR.WindowState = crptMaximized
    CR.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CR.Action = 1
    
    del_regi_qr.Execute

End Sub


