VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_collection_report 
   BackColor       =   &H8000000B&
   Caption         =   "SAMPLE COLLECTION REPORT"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14265
   LinkTopic       =   "Form2"
   ScaleHeight     =   8835
   ScaleWidth      =   14265
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   720
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ComboBox cmb_dept 
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
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   1200
      Width           =   4575
   End
   Begin VB.ComboBox Cmb_modality 
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
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox txt_tosrno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox txt_frmsrno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   3960
      Width           =   975
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
      Left            =   3000
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmd_print 
      BackColor       =   &H00C0E0FF&
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmd_BACK 
      BackColor       =   &H00C0E0FF&
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmd_clear 
      BackColor       =   &H00C0E0FF&
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   1215
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2760
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
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   1440
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lbl_dept 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1440
      TabIndex        =   17
      Top             =   1320
      Width           =   1230
   End
   Begin VB.Label lbl_mod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modality :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1680
      TabIndex        =   16
      Top             =   1920
      Width           =   930
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000B&
      Caption         =   "To Serial No :"
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
      Left            =   4680
      TabIndex        =   15
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000B&
      Caption         =   "From Serial No :"
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
      Left            =   1200
      TabIndex        =   14
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Collected Date : "
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
      Left            =   1200
      TabIndex        =   13
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000013&
      Height          =   5295
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   8055
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
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
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
      Left            =   1920
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8895
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   8520
      Y1              =   4680
      Y2              =   4680
   End
End
Attribute VB_Name = "frm_collection_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sel_location_qr As rdoQuery
Dim sel_location_rs As rdoResultset

Dim upd_TEMPCOLLCT_qr As rdoQuery

Dim SELRMODALITY_QR As rdoQuery
Dim SELRMODALITY_RS As rdoResultset

Dim var_date_f As String
Dim var_date_t As String
Dim MODCD, dept, stamp, MODALITY As String
Dim frsrno, tosrno, i, j As Integer


Private Sub cmb_dept_Click()
'Mrunal on 6_5_14
If cmb_dept.Text = "EMERGENCY LAB$FED" Then
'Label5.Enabled = False
'Label6.Enabled = False
'txt_frmsrno.Enabled = False
'txt_tosrno.Enabled = False
Else
'Label5.Enabled = True
'Label6.Enabled = True
'txt_frmsrno.Enabled = True
'txt_tosrno.Enabled = True

End If
'end
i = InStr(Trim(cmb_dept.Text), "$")
dept = Mid(cmb_dept.Text, i + 1, 4)
cmb_modality.clear
    If cmb_dept.Text = "EMERGENCY LAB$FED" Then
     cmb_modality.AddItem ("BIOCHEMISTRY$BC")
       cmb_modality.AddItem ("HAEMATOLOGY$HL")
     End If
    
    SELRMODALITY_QR(0) = Trim(dept)

    Set SELRMODALITY_RS = SELRMODALITY_QR.OpenResultset(1, 3)

If SELRMODALITY_RS.RowCount > 0 Then
    
    While Not SELRMODALITY_RS.EOF
        cmb_modality.AddItem SELRMODALITY_RS("modlitydes") & "$" & SELRMODALITY_RS("modalitycd")
        SELRMODALITY_RS.MoveNext
    Wend
End If
End Sub

Private Sub cmd_back_Click()
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_clear_Click()
MaskEdBox1.Text = "__/__/____"
'MaskEdBox2.Text = "__/__/____"
cmb_location.Text = ""
txt_frmsrno.Text = ""
txt_tosrno.Text = ""
MaskEdBox1.SetFocus
End Sub



Private Sub cmd_print_Click()

'***** Shraddha 02/08/2011 *****
i = InStr(Trim(cmb_dept.Text), "$")
dept = Mid(cmb_dept.Text, i + 1, 4)

If Trim(dept) = "" Then
    MsgBox "Please select depatment"
    Exit Sub
End If

MODCD = Right(Trim(cmb_modality.Text), 2)

If Trim(MODCD) = "" Then
    MsgBox "Please select modality"
    Exit Sub
End If
'*******************************

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

'If MaskEdBox2.Text = "__/__/____" Then
'    MsgBox "Please enter To date "
'    MaskEdBox2.Text = "__/__/____"
'    MaskEdBox2.SetFocus
'    Exit Sub
'ElseIf IsDate(MaskEdBox2.Text) = False Then
'    MsgBox "Please enter valid To date "
'    MaskEdBox2.Text = "__/__/____"
'    MaskEdBox2.SetFocus
'    Exit Sub
'End If


'Dim day_cnt
'day_cnt = DateDiff("d", MaskEdBox1.Text, MaskEdBox2.Text)
'If day_cnt < 0 Then
'   MsgBox "PLEASE ENTER VALID DATE RANGE"
'   MaskEdBox2.Text = "__/__/____"
'    MaskEdBox2.SetFocus
'    Exit Sub
'End If
'If cmb_dept.Text <> "EMERGENCY LAB$FED" Then mrunal
If Trim(txt_frmsrno.Text) = "" Then
    MsgBox "Please enter lab number"
    Exit Sub
End If

If Trim(txt_tosrno.Text) = "" Then
    MsgBox "Please enter lab number"
    Exit Sub
End If
'Else mrunal

'End If mrunal
'If Trim(cmb_location.Text) = "" Then
'    MsgBox "Please select location"
'    Exit Sub
'End If

Dim pop1 As Integer, pop2 As Integer

var_date_f = Year(MaskEdBox1.FormattedText) & "," & month(MaskEdBox1.FormattedText) & "," & Day(MaskEdBox1.FormattedText)
var_date_t = Year(MaskEdBox1.FormattedText) & "," & month(MaskEdBox1.FormattedText) & "," & Day(MaskEdBox1.FormattedText)
'var_date_t = Year(MaskEdBox2.FormattedText) & "," & month(MaskEdBox2.FormattedText) & "," & Day(MaskEdBox2.FormattedText)
    
frsrno = Trim(txt_frmsrno.Text)
tosrno = Trim(txt_tosrno.Text)
    
    


stamp = Right(UID, 6) & " - " & ipadd & " - " & Format(DATE11, "DD/MM/YYYY") & " - " & time1
    
j = InStr(cmb_modality.Text, "$")
MODALITY = Mid(cmb_modality.Text, 1, j - 1)
  '  If cmb_dept.Text <> "EMERGENCY LAB$FED" Then mrunal
    pop1 = frsrno
pop2 = tosrno
'CR.SelectionFormula = "{DMULPROCE.REQNDT}>=Date(" & var_date_f & ") AND {DMULPROCE.REQNDT}<=Date(" & var_date_t & ")"
'CrystalReport1.SelectionFormula = "{TEMPCOLLCT.CFLAG} = 'Y' AND {TEMPCOLLCT.SENTFLAG} = '' AND {TEMPCOLLCT.LOCATION}='" & Trim(cmb_location.Text) & "' AND {TEMPCOLLCT.MODALITYCD}='" & modcd & "'"
'CrystalReport1.SelectionFormula = "{TEMPCOLLCT.CFLAG} = 'Y' AND {TEMPCOLLCT.LOCATION}='" & Trim(cmb_location.Text) & "' AND {TEMPCOLLCT.MODALITYCD}='" & modcd & "' AND {TEMPCOLLCT.COLLDATE}>=Date(" & var_date_f & ") AND {TEMPCOLLCT.COLLDATE}<=Date(" & var_date_t & ") AND {TEMPCOLLCT.SRNO}>=" & frsrno & " AND {TEMPCOLLCT.SRNO}<= " & tosrno
'CrystalReport1.SelectionFormula = "{TEMPCOLLCT.CFLAG} = 'Y' AND {TEMPCOLLCT.LOCATION}='" & Trim(cmb_location.Text) & "' AND {TEMPCOLLCT.MODALITYCD}='" & MODCD & "' AND {TEMPCOLLCT.COLLDATE}>=Date(" & var_date_f & ") AND {TEMPCOLLCT.COLLDATE}<=Date(" & var_date_t & ") AND {TEMPCOLLCT.SRNO}>=" & frsrno & " AND {TEMPCOLLCT.SRNO}<= " & tosrno
'Commented and added by Shraddha on 13/10/2011

ST = Right(MaskEdBox1, 2) & Mid(MaskEdBox1, 4, 2) & Left(MaskEdBox1, 2)
If Len(Trim(txt_frmsrno.Text)) = 1 Then
    frsrno = "000" & Trim(txt_frmsrno.Text)
ElseIf Len(Trim(txt_frmsrno.Text)) = 2 Then
    frsrno = "00" & Trim(txt_frmsrno.Text)
ElseIf Len(Trim(txt_frmsrno.Text)) = 3 Then
    frsrno = "0" & Trim(txt_frmsrno.Text)
ElseIf Len(Trim(txt_frmsrno.Text)) = 4 Then
    frsrno = Trim(txt_frmsrno.Text)
End If

If Len(Trim(txt_tosrno.Text)) = 1 Then
    tosrno = "000" & Trim(txt_tosrno.Text)
ElseIf Len(Trim(txt_tosrno.Text)) = 2 Then
    tosrno = "00" & Trim(txt_tosrno.Text)
ElseIf Len(Trim(txt_tosrno.Text)) = 3 Then
    tosrno = "0" & Trim(txt_tosrno.Text)
ElseIf Len(Trim(txt_tosrno.Text)) = 4 Then
    tosrno = Trim(txt_tosrno.Text)
End If

frsrno = ST & frsrno

tosrno = ST & tosrno

'End If mrunal



           
'CrystalReport1.SelectionFormula = "{TEMPCOLLCT.CFLAG} = 'Y' AND {TEMPCOLLCT.MODALITYCD}='" & MODCD & "' AND ToNumber({DREQUISITN.PATHNO})>=" & frsrno & " AND Tonumber({DREQUISITN.PATHNO})<= " & tosrno
If cmb_dept.Text <> "EMERGENCY LAB$FED" Then
CrystalReport1.SelectionFormula = "{TEMPCOLLCT.CFLAG} = 'Y' AND {TEMPCOLLCT.MODALITYCD}='" & MODCD & "' AND {DREQUISITN.PATHNO} in '" & frsrno & "' to '" & tosrno & "'"
Else
CrystalReport1.SelectionFormula = "{TEMPCOLLCT.CFLAG} = 'Y' AND {TEMPCOLLCT.MODALITYCD}='" & MODCD & "' and {TEMPCOLLCT.SENTDATE} =Date (" & var_date_f & ")  AND {DREQUISITN.PATHNO} in '" & frsrno & "' to  '" & tosrno & "' "
'and {TEMPCOLLCT.SENTDATE}>= Date(" & c_dt1 & ") AND {TEMPCOLLCT.SENTDATE} <= Date(" & dt4 & ")"

End If

'NN = Right(txt_frmsrno, 2) & Mid(msk_fromdt.FormattedText, 4, 2) & Left(msk_fromdt.FormattedText, 2)
'txt_frmLabNo.Text = NN & txt_frmLabNo.Text
'***** Shraddha 16/11/2011 *****
If cmb_dept.Text <> "EMERGENCY LAB$FED" Then
If MODCD = "BC" And CInt(Trim(txt_frmsrno.Text)) >= 601 And CInt(Trim(txt_tosrno.Text)) <= 700 Then
    CrystalReport2.ReportFileName = "C:\DISREPORTSEC\LISTSTAFFBC.rpt"
    CrystalReport2.SelectionFormula = "{LISTSTAFFBC.PATHNO} in '" & frsrno & "' to '" & tosrno & "'"
    'CrystalReport2.SelectionFormula = "{LISTSTAFFBC.PATHNO} in "" & frsrno & "" to "" & tosrno &"""
    CrystalReport2.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
    CrystalReport2.WindowState = crptMaximized
    CrystalReport2.Action = 1
    Exit Sub
ElseIf MODCD = "HL" Then
    
stamp = Right(UID, 6) & " - " & ipadd & " - " & Format(DATE11, "DD/MM/YYYY") & " - " & time1
CrystalReport1.SelectionFormula = "{TEMPRECEIV.MODALITYCD} = 'HL' AND  {TEMPRECEIV.SRNO} in " & CInt(Trim(txt_frmsrno.Text)) & " to " & CInt(Trim(txt_tosrno.Text)) & " AND  {TEMPRECEIV.PRLSTDT}= CurrentDate"
CrystalReport1.Formulas(0) = "STAMP = '" & stamp & "'"
CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\SampleRcvHL.RPT"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport1.Action = 1


Exit Sub
End If
End If
'*******************************
Dim pop As String
If pop1 >= 1 And pop2 < 400 Then
pop = "WARD"
ElseIf pop1 >= 401 And pop2 < 480 Then
pop = "PROTOCOL / RF / ACTREC"
ElseIf pop1 >= 501 And pop2 < 600 Then
pop = "URGENT [ICU 1ST & 2ND FLOOR "
ElseIf pop1 >= 601 And pop2 < 700 Then
pop = "STAFF"
ElseIf pop1 >= 701 And pop2 < 1500 Then
pop = "PVT OPD"
ElseIf pop1 >= 1501 Then
pop = "GENERAL OPD"
End If


CrystalReport1.Formulas(0) = "STAMP = '" & stamp & "'"
CrystalReport1.Formulas(1) = "MODALITY = '" & MODALITY & "'"
CrystalReport1.Formulas(2) = "POP = '" & Trim(pop) & "'"
If cmb_dept.Text <> "EMERGENCY LAB$FED" Then
CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\SampleCollection.RPT"
Else
 CrystalReport1.ReportFileName = "C:\DISREPORTSECLF\SampleCollection_fed.rpt"
End If
CrystalReport1.WindowState = crptMaximized
CrystalReport1.Connect = "ODBC;Database=DISLIB;UID=" & Trim(LOGIN) & ";PWD= " & Trim(LOGIN) & ";DSN=DISDSN"
CrystalReport1.Action = 1



'upd_TEMPCOLLCT_qr.Execute

End Sub

Private Sub Form_Load()

str = "update Dislib.TEMPCOLLCT set SENTFLAG='P' WHERE SENTFLAG='' AND CFLAG='Y'"
Set upd_TEMPCOLLCT_qr = rdb.CreateQuery("", str)

str = "select * from patoolslib.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ'  OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ' or dept_code='IZZ' or dept_code='CCH' or dept_code='FED')"
Set SELMA106_QR = rdb.CreateQuery("", str)

str = "select * from rislib.rmodality where dept_code=?"
Set SELRMODALITY_QR = rdb.CreateQuery("", str)

str = "select distinct(location) from RISLIB.rlocation ORDER BY LOCATION"
Set sel_location_qr = rdb.CreateQuery("", str)
Set sel_location_rs = sel_location_qr.OpenResultset(1, 3)

While Not sel_location_rs.EOF
    cmb_location.AddItem Trim(sel_location_rs("LOCATION"))
    sel_location_rs.MoveNext
Wend

datetime

MaskEdBox1.Text = Format(DATE11, "dd/mm/yyyy")
'MaskEdBox2.Text = Format(DATE11, "dd/mm/yyyy")

Set SELMA106_RS = SELMA106_QR.OpenResultset(1, 3)

If SELMA106_RS.RowCount > 0 Then
    cmb_dept.clear
    While Not SELMA106_RS.EOF
        If Trim(USRDEPT) <> "ALL" Then
            If Trim(USRDEPT) = Trim(SELMA106_RS("DEPT_CODE")) Then
                cmb_dept.AddItem Trim(SELMA106_RS("DEPT_NAME")) & "$" & Trim(SELMA106_RS("DEPT_CODE"))
            End If
        Else
             cmb_dept.AddItem Trim(SELMA106_RS("DEPT_NAME")) & "$" & Trim(SELMA106_RS("DEPT_CODE"))
        End If
        SELMA106_RS.MoveNext
    Wend
End If

End Sub

