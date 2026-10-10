VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form regi 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
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
      Left            =   2760
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
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
      Left            =   720
      TabIndex        =   3
      Top             =   2640
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
      Left            =   5760
      TabIndex        =   2
      Top             =   2640
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
      Left            =   4080
      TabIndex        =   1
      Top             =   2640
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
      Left            =   2520
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1080
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
      Left            =   5760
      TabIndex        =   6
      Top             =   1080
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
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
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
      Left            =   600
      TabIndex        =   10
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000013&
      Height          =   2535
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   7095
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
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
      Left            =   4800
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
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
      Left            =   1680
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000013&
      Caption         =   "TATA MEMORIAL HOSPITAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   4215
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   7440
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "regi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
' This form has been developed by Ms. Shraddha Pandit on 04/05/2009
'*******************************************************************

Dim modality As String
Dim dept As String
Dim var_date_f As String
Dim var_date_t As String

Dim sql As String
Dim sql1 As String
Dim sql2 As String
Dim sql3 As String
Dim sql4 As String

Dim sel_loc_qr As rdoQuery
Dim sel_loc_rs As rdoResultset

Dim sel_rtimelog_qr As rdoQuery
Dim sel_rtimelog_rs As rdoResultset

Dim SEL_PROCODE_QR As rdoQuery
Dim SEL_PROCODE_RS As rdoResultset

Dim INS_REGI_QR As rdoQuery

Dim del_regi_qr As rdoQuery

Private Sub Form_Load()

sql1 = "SELECT DISTINCT REQNNO, CASENO, SAMCOLDATE FROM rislib.rtimelog WHERE samcoldate >= ? AND samcoldate <= ?"
Set sel_rtimelog_qr = rdb.CreateQuery("", sql1)

sql2 = "DELETE FROM DISLIB.registerop"
Set del_regi_qr = rdb.CreateQuery("", sql2)

sql3 = "INSERT INTO DISLIB.registerop (CASENO, REQNO, LOCATION, COLLECTDT, BIOCHEM, TUMOUR, CBC, COAG_PROF, SCREENING, MICROBIOLO) VALUES(?,?,?,?,?,?,?,?,?,?)"
Set INS_REGI_QR = rdb.CreateQuery("", sql3)

sql4 = "SELECT PRO_CODE FROM DISLIB.DMULPROCE WHERE REQNO=?"
Set SEL_PROCODE_QR = rdb.CreateQuery("", sql4)

sql = "select distinct(location) from dislib.barcode"
Set sel_loc_qr = rdb.CreateQuery("", sql)

Set sel_loc_rs = sel_loc_qr.OpenResultset(1, 3)
While Not sel_loc_rs.EOF
               
    If Trim(sel_loc_rs(0)) <> "" Then
        cmb_location.AddItem Trim(sel_loc_rs(0))
    End If
    sel_loc_rs.MoveNext
                
Wend


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

    var_date_f = Year(MaskEdBox1.FormattedText) & "," & Month(MaskEdBox1.FormattedText) & "," & Day(MaskEdBox1.FormattedText)
    var_date_t = Year(MaskEdBox2.FormattedText) & "," & Month(MaskEdBox2.FormattedText) & "," & Day(MaskEdBox2.FormattedText)
   
    sel_rtimelog_qr(0) = var_date_f
    sel_rtimelog_qr(1) = var_date_t

    Set sel_rtimelog_rs = sel_rtimelog_qr.OpenResultset(1, 3)

    While Not sel_rtimelog_rs.EOF
    
        dept = Mid(sel_rtimelog_rs("REQNNO"), 1, 3)
        modality = Mid(sel_rtimelog_rs("REQNNO"), 4, 2)
        
        If dept = "FEC" Or dept = "UZZ" Or dept = "NZZ" Or dept = "GZZ" Then
        
            INS_REGI_QR(0) = Trim(sel_rtimelog_rs("CASENO"))
            INS_REGI_QR(1) = Trim(sel_rtimelog_rs("REQNNO"))
            INS_REGI_QR(2) = cmb_location.Text
            INS_REGI_QR(3) = sel_rtimelog_rs("SAMCOLDATE")
            
            If dept = "FEC" Then
                INS_REGI_QR(4) = "YES"
            Else
                INS_REGI_QR(4) = ""
            End If
            
            If modality = "TM" Then
                INS_REGI_QR(5) = "YES"
            Else
                INS_REGI_QR(5) = ""
            End If
    
            'TO SELECT PRO_CODE
            SEL_PROCODE_QR(0) = sel_rtimelog_rs("REQNNO")
            Set SEL_PROCODE_RS = SEL_PROCODE_QR.OpenResultset(1, 3)
    
            While Not SEL_PROCODE_RS.EOF
    
                If SEL_PROCODE_RS(0) = "U721" Then
                    INS_REGI_QR(6) = "Yes" 'CBC
                Else
                    INS_REGI_QR(6) = ""
                End If
    
                If SEL_PROCODE_RS(0) = "U709" Then
                    INS_REGI_QR(7) = "Yes" 'Coag_prof
                Else
                    INS_REGI_QR(7) = ""
                End If
    
                If SEL_PROCODE_RS(0) = "U706" Then
                    INS_REGI_QR(8) = "Yes" 'ESR
                Else
                    INS_REGI_QR(8) = ""
                End If
    
                SEL_PROCODE_RS.MoveNext
            Wend
    
            If dept = "NZZ" And modality = "SS" Then
                INS_REGI_QR(9) = "YES" 'screening
            Else
                INS_REGI_QR(9) = ""
            End If
            
            If dept = "GZZ" And modality = "SE" Then
                INS_REGI_QR(9) = "YES" 'microbio
            Else
                INS_REGI_QR(9) = ""
            End If
    
            INS_REGI_QR.Execute
            
        End If
        
        sel_rtimelog_rs.MoveNext
        
    Wend
    

 sel_loc_rs.Close


 sel_rtimelog_rs.Close


 SEL_PROCODE_RS.Close
    

    CR.ReportFileName = "C:\DISREPORTSEC\SampleDaily_Dispatch.rpt"
    CR.Formulas(0) = "from=Date(" & var_date_f & ")"
    CR.Formulas(1) = "to=Date(" & var_date_t & ")"
    'CR.SelectionFormula = "{DMULPROCE.REQNDT}>=Date(" & var_date_f & ") AND {DMULPROCE.REQNDT}<=Date(" & var_date_t & ")"
    CR.WindowState = crptMaximized
    CR.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CR.Action = 1

    del_regi_qr.Execute
  

End Sub



