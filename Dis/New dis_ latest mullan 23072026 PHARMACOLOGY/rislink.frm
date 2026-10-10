VERSION 5.00
Begin VB.Form RISLINK 
   Caption         =   "Instruction Group"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   6840
      Width           =   11775
      Begin VB.CommandButton cmdnew 
         Caption         =   "&NEW"
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
         Left            =   1560
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdfirst 
         Caption         =   "&FIRST"
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
         Left            =   2550
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdlast 
         Caption         =   "&LAST"
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
         Left            =   5820
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&NEXT"
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
         Left            =   4860
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&PREVIOUS"
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
         Left            =   3630
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdmodify 
         Caption         =   "&MODIFY"
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
         Left            =   6780
         TabIndex        =   16
         Top             =   360
         Width           =   1092
      End
      Begin VB.CommandButton cmddelte 
         Caption         =   "&DELETE"
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
         Left            =   7860
         TabIndex        =   15
         Top             =   360
         Width           =   972
      End
      Begin VB.CommandButton cmdexit 
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
         Left            =   9900
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdcancel1 
         Caption         =   "C&ANCEL"
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
         Left            =   8820
         TabIndex        =   13
         Top             =   360
         Width           =   1092
      End
      Begin VB.CommandButton cmdview 
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
         Left            =   3840
         TabIndex        =   24
         Top             =   360
         Width           =   1092
      End
      Begin VB.CommandButton cmdcancel 
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
         Left            =   2550
         TabIndex        =   25
         Top             =   360
         Width           =   1035
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&SAVE"
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
         Left            =   1560
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2550
         TabIndex        =   23
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label lblmesg 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3720
         TabIndex        =   22
         Top             =   960
         Width           =   6015
      End
   End
   Begin VB.TextBox TXT_INS 
      Appearance      =   0  'Flat
      Height          =   2505
      Left            =   900
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3480
      Width           =   10125
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      TabIndex        =   7
      Top             =   1515
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9840
      TabIndex        =   6
      Top             =   1440
      Width           =   1035
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4620
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2820
      Width           =   855
   End
   Begin VB.TextBox txtinsno 
      Height          =   285
      Left            =   4620
      MaxLength       =   3
      TabIndex        =   1
      Top             =   2820
      Width           =   855
   End
   Begin VB.TextBox txtgrpno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4620
      MaxLength       =   3
      TabIndex        =   0
      Top             =   2100
      Width           =   855
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
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   240
      Width           =   11655
   End
   Begin VB.Label LBL_INS 
      Caption         =   "INSTRUCTION TEXT"
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
      Left            =   930
      TabIndex        =   11
      Top             =   3210
      Width           =   1845
   End
   Begin VB.Label Label5 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   8
      Top             =   1470
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   " INSTRUCTION GROUP"
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
      Height          =   330
      Left            =   3990
      TabIndex        =   5
      Top             =   690
      Width           =   3210
   End
   Begin VB.Label lblinstest 
      Caption         =   "INSTRUCTION NO."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2790
      TabIndex        =   4
      Top             =   2850
      Width           =   1695
   End
   Begin VB.Label lblinsno 
      Caption         =   "GROUP NO."
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
      Left            =   3420
      TabIndex        =   3
      Top             =   2130
      Width           =   1065
   End
End
Attribute VB_Name = "RISLINK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsq As rdoQuery
Dim rsqq As rdoQuery
Dim RS As rdoResultset
Dim rs2 As rdoResultset
Dim rs1 As rdoResultset
Dim SELINST_QR As rdoQuery
Dim SELINST_RS As rdoResultset
Dim selrislnklf_qr As rdoQuery
Dim selrislnklf_rs As rdoResultset
Dim STRSQL As String
Dim strsql1 As String
Dim opmode As String
Dim savemode As String







Private Sub cmdCancel_Click()
cmdnew.Visible = True
cmddelte.Visible = False
cmdexit.Visible = True

cmdfirst.Visible = False
cmdprev.Visible = False
cmdlast.Visible = False
cmdnext.Visible = False
cmdmodify.Visible = False
cmdsave.Visible = False
cmdcancel.Visible = False
cmdview.Visible = True
'txtgrpno.Text = ""
'TXT_INS.Text = ""
'txtinsno.Text = ""
End Sub

Private Sub cmdcancel1_Click()
cmdnew.Visible = True
cmddelte.Visible = False
cmdexit.Visible = True

cmdfirst.Visible = False
cmdprev.Visible = False
cmdlast.Visible = False
cmdnext.Visible = False
cmdmodify.Visible = False
cmdsave.Visible = False
cmdcancel.Visible = False
cmdcancel1.Visible = False
cmdview.Visible = True
txtgrpno.Text = ""
TXT_INS.Text = ""
txtinsno.Text = ""
cmdnew.SetFocus
End Sub

Private Sub cmddelte_Click()
lblmesg = " "
On Error GoTo MSG

If Len(Trim((txtinsno.Text))) = 0 Then
 MsgBox "Nothing selected to delete"
  Exit Sub
  
End If
'If (RS.EOF) Or (RS.BOF) Then
'    MsgBox " Fisrt/Last record is crossed cannot delete"
' Exit Sub
'Else

      RS.MoveFirst
   While Not RS.EOF
     If Trim(RS("insno")) = CInt(Trim(txtinsno)) Then
          
        RS.Delete
        lblmesg = "Record Deleted "
        Exit Sub
     End If
    RS.MoveNext
   Wend
   
   
   
'End If
MSG:
MsgBox Err.Number

  
   
End Sub

Private Sub cmdexit_Click()


DIS_MAIN.show
Me.Hide

End Sub

Private Sub cmdfirst_Click()
Combo1.Visible = False
txtinsno.Visible = True
  selrislnklf_rs.MoveFirst
  'txtgrpno = rs("grpno")
  txtinsno = selrislnklf_rs("insno")
  TXT_INS.Text = selrislnklf_rs("instext")
End Sub

Private Sub cmdlast_Click()
Combo1.Visible = False
txtinsno.Visible = True
selrislnklf_rs.MoveLast
  'txtgrpno = rs("grpno")
  txtinsno = selrislnklf_rs("insno")
  TXT_INS.Text = selrislnklf_rs("instext")
End Sub

Private Sub cmdmodify_Click()
On Error GoTo MSG

If Len(Trim((txtgrpno.Text))) = 0 Then
 MsgBox "Nothing selected to modify"
  Exit Sub
  
End If
'If (rs.EOF) Or (rs.BOF) Then
'    MsgBox " Fisrt/Last record is crossed cannot modify"
' Exit Sub
'
'End If
savemode = "M"
cmdnew.Visible = False
cmddelte.Visible = False
cmdexit.Visible = False
cmdfirst.Visible = False
cmdprev.Visible = False
cmdlast.Visible = False
cmdnext.Visible = False
cmdmodify.Visible = False
cmdsave.Visible = True
cmdcancel.Visible = True
cmdview.Visible = False
cmdcancel1.Visible = False
txtinsno.Visible = False
Combo1.Visible = True
Combo1.Text = txtinsno.Text

'txtinsno = " "
'txtinstext = " "
'txtinsno.SetFocus
Exit Sub

MSG:
MsgBox Err.Number

End Sub

Private Sub cmdnew_Click()
opmode = "N"
savemode = "N"
cmdnew.Visible = False
cmddelte.Visible = False
cmdexit.Visible = False
cmdfirst.Visible = False
cmdprev.Visible = False
cmdlast.Visible = False
cmdnext.Visible = False
cmdmodify.Visible = False
cmdsave.Visible = True
cmdcancel.Visible = True
cmdview.Visible = False
cmdcancel1.Visible = False
txtinsno.Visible = False
Combo1.Visible = True

txtgrpno = " "
txtinsno = " "
'Combo1.Text = " "
Combo1.ListIndex = -1

txtgrpno.SetFocus
Text1.Text = time1
Text2.Text = DATE11

End Sub

Private Sub cmdnext_Click()
On Error GoTo MSG:
Combo1.Visible = False
txtinsno.Visible = True
selrislnklf_rs.MoveNext
 If selrislnklf_rs.EOF Then
     MsgBox "Last Record"
     cmdnext.SetFocus
 Else
   
   'txtgrpno = rs("grpno")
  txtinsno = selrislnklf_rs("insno")
  TXT_INS.Text = selrislnklf_rs("instext")
 End If
''rs.MoveNext
'' If rs.EOF Then
''     MsgBox "Last Record"
''     cmdnext.SetFocus
'' Else
''
''   txtgrpno = rs("grpno")
''  txtinsno = rs("insno")
'' End If
 Exit Sub
 
MSG:
   MsgBox Err.Description
   cmdnext.SetFocus
   
 
End Sub

Private Sub cmdprev_Click()
On Error GoTo MSG:
Combo1.Visible = False
txtinsno.Visible = True
selrislnklf_rs.MovePrevious

If selrislnklf_rs.BOF Then
     MsgBox "First Record"
     cmdprev.SetFocus
 Else

  ' txtgrpno = rs("grpno")
  txtinsno = selrislnklf_rs("insno")
  TXT_INS.Text = selrislnklf_rs("instext")
 End If



' If rs.BOF Then
'     MsgBox "First Record"
'     cmdprev.SetFocus
' Else
'
'   txtgrpno = rs("grpno")
'  txtinsno = rs("insno")
' End If
 Exit Sub
 
MSG:
  MsgBox Err.Description
  cmdprev.SetFocus
'  Exit Sub
  
End Sub

Private Sub cmdSave_Click()
lblmesg = " "
On Error GoTo MSG:
If savemode = "N" Then

 RS.AddNew
 RS("grpno") = txtgrpno
 RS("insno") = Combo1.Text
 'rs.Edit
 RS.Update
 lblmesg = "Record Added "
End If
If savemode = "M" Then

      RS.MoveFirst
   While Not RS.EOF
     If CInt((Trim(RS("grpno"))) = CInt(Trim(txtgrpno))) And CInt((Trim(RS("insno"))) = CInt(Trim(txtinsno))) Then
     
         ' fndflg = True
      'txtinstext = rs("instext")
      RS.Edit
      RS("grpno") = txtgrpno
      RS("insno") = Combo1.Text
       'rs.Edit
      RS.Update
      lblmesg = "Record Modified  "
     End If
    RS.MoveNext
   Wend

''
 
End If
 If savemode = "N" Then
     txtgrpno = " "
    txtinsno = " "
 End If
  txtgrpno.SetFocus
 Exit Sub
 
MSG:
If Err.Number = 40002 Then
    MsgBox "Duplicate Record Error"
    txtgrpno = " "
    txtinsno = " "
    RS.Requery
    txtgrpno.SetFocus
    Exit Sub
End If
If Err.Number <> 0 Then

  MsgBox Err.Number
  MsgBox Err.Description
  RS.Requery
End If

 
   
    
End Sub

Private Sub CMDVIEW_Click()
cmdnew.Visible = False
cmddelte.Visible = False
cmdexit.Visible = False
cmdfirst.Visible = True
cmdprev.Visible = True
cmdlast.Visible = True
cmdnext.Visible = True
cmdmodify.Visible = True
cmdsave.Visible = False
cmdcancel.Visible = False
cmdcancel1.Visible = True
cmdview.Visible = False
cmddelte.Visible = True
Set RS = rsq.OpenResultset(2, 3)
txtgrpno = " "
txtinsno = " "
opmode = "V"
Combo1.Visible = False
txtinsno.Visible = True
txtgrpno.SetFocus
Text1.Text = time1
Text2.Text = DATE11
LBL_INS.Visible = True
TXT_INS.Visible = True
End Sub

Private Sub Combo1_Click()
If Trim(Combo1.Text) = "" Then
Exit Sub
End If
SELINST_QR(0) = Trim(Combo1.Text)
SELINST_RS.MoreResults
SELINST_RS.Requery
If SELINST_RS.RowCount > 0 Then
TXT_INS.Text = SELINST_RS("INSTEXT")
End If

End Sub

Private Sub Form_Load()
'Main
Set rs2 = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107") '***2/2/99
DATE11 = rs2("dt1") '**add 24/11/2k
time1 = Format(rs2("TM1"), "hh:mm:ss") '***2/2/99
cmdnew.Visible = True
cmddelte.Visible = False
cmdexit.Visible = False
cmdfirst.Visible = False
cmdprev.Visible = False
cmdlast.Visible = False
cmdnext.Visible = False
cmdmodify.Visible = False
cmdsave.Visible = False
cmdcancel.Visible = False
cmdcancel1.Visible = False
cmdview.Visible = True
cmdexit.Visible = True
Combo1.Visible = False
opmode = " "
savemode = " "
lblmesg = " "
STRSQL = "select * from rislib.rislink order by grpno,insno"
Set rsq = rdb.CreateQuery("qr11", STRSQL)
Set RS = rsq.OpenResultset(2, 3)
strsql1 = "select * from rislib.risinstruc order by insno"
Set rsqq = rdb.CreateQuery("rsqq1", strsql1)
Set rs1 = rsqq.OpenResultset(1, 1)
Combo1.clear
If rs1.RowCount > 0 Then

 rs1.MoveFirst
 While Not rs1.EOF
  Combo1.AddItem rs1("insno")
  rs1.MoveNext
 Wend
End If
Set SELINST_QR = rdb.CreateQuery("SELINST_QR1", "SELECT * FROM RISLIB.RISINSTRUC WHERE INSNO=?")
SELINST_QR(0) = 0
Set SELINST_RS = SELINST_QR.OpenResultset(2, 3)
Set selrislnklf_qr = rdb.CreateQuery("selrislnklf_qr1", "select * from rislib.rislnklf where grpno=?")
selrislnklf_qr(0) = 0
Set selrislnklf_rs = selrislnklf_qr.OpenResultset(2, 3)
Text1.Text = time1
Text2.Text = DATE11
'LBL_INS.Visible = False
'TXT_INS.Visible = False

End Sub

Private Sub txtgrpnoi_Change()



End Sub

Private Sub Form_Unload(cancel As Integer)
Dim RS As rdoResultset
Dim RQ As rdoQuery
rsq.Close
rsqq.Close
For Each RS In rdb.rdoResultsets
    RS.Close
Next

For Each RQ In rdb.rdoQueries
    RQ.Close
Next
End Sub

Private Sub txtgrpno_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48) Or (KeyAscii > 57) Then
  KeyAscii = 0
End If

End Sub

Private Sub txtgrpno_LostFocus()
'Dim fndflg As Boolean
' If opmode = "V" Then
'  If Len(Trim((txtgrpno.Text))) > 0 Then
'
'   fndflg = False
'
'      rs.MoveFirst
'   While Not rs.EOF
'     If Trim(rs("grpno")) = CInt(Trim(txtgrpno)) Then
'          fndflg = True
'      txtinsno = rs("insno")
'     End If
'    rs.MoveNext
'   Wend
'   If fndflg = False Then
'        MsgBox "Record Not Found"
'        txtgrpno.SetFocus
'   End If
'
'
' End If
'End If
If Trim(txtgrpno.Text) = "" Then
txtgrpno.Text = ""
TXT_INS.Text = ""
txtinsno.Text = ""
Exit Sub
End If
selrislnklf_qr(0) = Trim(txtgrpno.Text)
selrislnklf_rs.MoreResults
selrislnklf_rs.Requery
If selrislnklf_rs.RowCount > 0 Then
selrislnklf_rs.MoveFirst
txtinsno.Text = selrislnklf_rs("insno")
TXT_INS.Text = selrislnklf_rs("instext")
Else
MsgBox "No records are there"
Exit Sub
End If

End Sub

Private Sub txtinsno_LostFocus()
Dim fndflg As Boolean
 If opmode = "V" Then
  If (Len(Trim((txtgrpno.Text)))) = 0 And (Len(Trim((txtinsno.Text)))) = 0 Then
    MsgBox " BOth Group No. and Instruction No. not entered"
    Exit Sub
  End If
  
  If (Len(Trim((txtgrpno.Text)))) > 0 And (Len(Trim((txtinsno.Text)))) > 0 Then

   fndflg = False

      RS.MoveFirst
   While Not RS.EOF
     If CInt((Trim(RS("grpno"))) = CInt((Trim(txtgrpno)))) And CInt((Trim(RS("insno"))) = CInt((Trim(txtinsno)))) Then
          fndflg = True
      txtgrpno = RS("grpno")
      txtinsno = RS("insno")
     End If
    RS.MoveNext
   Wend
   If fndflg = False Then
        MsgBox "Record Not Found"
        txtgrpno.SetFocus
   End If


 End If
End If

End Sub
