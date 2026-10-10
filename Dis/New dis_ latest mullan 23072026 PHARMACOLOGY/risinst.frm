VERSION 5.00
Begin VB.Form RISINST 
   Caption         =   "Instruction Master"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   7200
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
         Left            =   120
         TabIndex        =   18
         Top             =   240
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
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   975
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
         Left            =   4350
         TabIndex        =   16
         Top             =   240
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
         Left            =   3390
         TabIndex        =   15
         Top             =   240
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
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   1332
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
         Left            =   5310
         TabIndex        =   13
         Top             =   240
         Width           =   972
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
         Left            =   6270
         TabIndex        =   12
         Top             =   240
         Width           =   1092
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
         Left            =   8460
         TabIndex        =   11
         Top             =   240
         Width           =   1095
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
         Left            =   7350
         TabIndex        =   10
         Top             =   240
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
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   1212
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
         Left            =   2040
         TabIndex        =   22
         Top             =   240
         Width           =   1335
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
         Left            =   120
         TabIndex        =   23
         Top             =   240
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
         Left            =   660
         TabIndex        =   20
         Top             =   810
         Width           =   1125
      End
      Begin VB.Label lblmesg 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1830
         TabIndex        =   19
         Top             =   780
         Width           =   7365
      End
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
      Height          =   300
      Left            =   2430
      TabIndex        =   6
      Top             =   2355
      Width           =   1245
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
      Height          =   285
      Left            =   7470
      TabIndex        =   5
      Top             =   2370
      Width           =   1035
   End
   Begin VB.TextBox txtinstext 
      Appearance      =   0  'Flat
      Height          =   3132
      Left            =   330
      MaxLength       =   801
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3630
      Width           =   11145
   End
   Begin VB.TextBox txtinsno 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2430
      MaxLength       =   3
      TabIndex        =   1
      Top             =   2700
      Width           =   1245
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
      TabIndex        =   24
      Top             =   240
      Width           =   11655
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
      Left            =   1890
      TabIndex        =   8
      Top             =   2400
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
      Left            =   6810
      TabIndex        =   7
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "INSTRUCTIONS"
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
      Height          =   345
      Left            =   4560
      TabIndex        =   4
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lblinstest 
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
      Height          =   225
      Left            =   4530
      TabIndex        =   2
      Top             =   3330
      Width           =   1935
   End
   Begin VB.Label lblinsno 
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
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "RISINST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsq As rdoQuery
Dim RS As rdoResultset
Dim STRSQL As String
Dim opmode As String
Dim savemode As String
Dim rs1 As rdoResultset
Dim NAVMODE As String
Dim i As Integer





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
txtinsno.Text = ""
txtinstext.Text = ""
End Sub

Private Sub cmddelte_Click()
lblmesg = " "
On Error GoTo MSG

If Len(Trim((txtinsno.Text))) = 0 Then
 MsgBox "There is NOTHING to DELETE."
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
  RS.MoveFirst
  txtinsno = RS("insno")
  txtinstext = RS("instext")
  
End Sub

Private Sub cmdlast_Click()
RS.MoveLast
  txtinsno = RS("insno")
  txtinstext = RS("instext")
End Sub

Private Sub cmdmodify_Click()
On Error GoTo MSG

If Len(Trim((txtinsno.Text))) = 0 Then
 MsgBox "There is NOTHING to MODIFY."
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
'txtinsno = " "
'txtinstext = " "
'txtinsno.SetFocus
Exit Sub

MSG:
MsgBox Err.Number

End Sub

Private Sub cmdnew_Click()
i = 0
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
txtinsno = " "
txtinstext = " "
txtinsno.SetFocus
Text1.Text = time1
Text2.Text = DATE11
End Sub

Private Sub cmdnext_Click()
On Error GoTo MSG:
If NAVMODE = "V" Then
'RS.MoveNext
' If RS.EOF Then
'     MsgBox "This is the LAST RECORD."
'     cmdnext.SetFocus
'     Exit Sub
'
' Else
'   txtinsno = RS("insno")
'  txtinstext = RS("instext")
'  NAVMODE = "F"
'  GoTo L1:
' End If
' Exit Sub
RS.MoveFirst
While Not RS.EOF
     If Trim(RS("INSNO")) = (Trim(txtinsno)) Then
          'fndflg = True
          'MsgBox "HERE FOUND"
          
      'txtinstext = Trim(RS("USERNAME"))
      NAVMODE = "F"
      GoTo L1:
      
      Else
        RS.MoveNext
     End If
   Wend




End If
L1:
   RS.MoveNext
   If RS.EOF Then
     MsgBox "This is the LAST RECORD."
     cmdnext.SetFocus
     Exit Sub
   Else
    txtinsno = RS("insno")
   txtinstext = RS("instext")
    End If
    Exit Sub
MSG:

 MsgBox Err.Description
 
 
End Sub

Private Sub cmdprev_Click()
On Error GoTo MSG:


RS.MovePrevious

 If RS.BOF Then
     MsgBox "This is the FIRST RECORD."
     cmdprev.SetFocus
     Exit Sub
     
 Else
   
   txtinsno = RS("insno")
  txtinstext = RS("instext")
 End If
 Exit Sub
MSG:
   MsgBox Err.Description
   
End Sub

Private Sub cmdSave_Click()
lblmesg = " "

On Error GoTo MSG:

 While InStr(txtinstext, Chr(9)) > 0
            Dim beforst As String, afterst As String
            Dim K As Long
            K = InStr(txtinstext, Chr(9))
            beforst = Mid(txtinstext, 1, K - 1) & "      "
            afterst = Mid(txtinstext, K + 1)
            txtinstext = beforst & afterst
       Wend
While InStr(txtinstext, Chr(146)) > 0
'            Dim beforst As String, afterst As String
'            Dim k As Long
            K = InStr(txtinstext, Chr(146))
            beforst = Mid(txtinstext, 1, K - 1) & "`"
            afterst = Mid(txtinstext, K + 1)
            txtinstext = beforst & afterst
       Wend
While InStr(txtinstext, Chr(13)) > 0
'            Dim beforst As String, afterst As String
'            Dim k As Long
            K = InStr(txtinstext, Chr(13))
            beforst = Mid(txtinstext, 1, K - 1) & "  "
            afterst = Mid(txtinstext, K + 1)
            txtinstext = beforst & afterst
       Wend




If savemode = "N" Then

 RS.AddNew
 RS("insno") = txtinsno
 RS("instext") = Null
 
 RS("instext") = txtinstext
 'rs.Edit
 RS.Update
 lblmesg = "Record Added "
End If
If savemode = "M" Then

      RS.MoveFirst
   While Not RS.EOF
     If Trim(RS("insno")) = CInt(Trim(txtinsno)) Then
         ' fndflg = True
      'txtinstext = rs("instext")
      RS.Edit
      RS("insno") = txtinsno
       RS("instext") = Null
      RS("instext") = txtinstext
       'rs.Edit
      RS.Update
      lblmesg = "Record Modified  "
     End If
    RS.MoveNext
   Wend

''
 
End If
 
MSG:
If Err.Number = 40002 Then
    MsgBox "There is a Duplicate Record ERROR."
    txtinsno = " "
    txtinstext = " "
    RS.Requery
    txtinsno.SetFocus
    Exit Sub
End If
If Err.Number <> 0 Then

  MsgBox Err.Number
  MsgBox Err.Description
  RS.Requery
End If
 If savemode = "N" Then
     txtinsno = " "
    txtinstext = " "
 End If
 
    txtinsno.SetFocus
    
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
'rs.MoreResults
'rs.Requery
Set RS = rsq.OpenResultset(2, 3)
txtinsno = " "
txtinstext = " "
opmode = "V"
Text1.Text = time1
Text2.Text = DATE11
txtinsno.SetFocus

End Sub

Private Sub Form_Load()
'Main
Set rs1 = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107") '***2/2/99
DATE11 = rs1("dt1") '**add 24/11/2k
time1 = Format(rs1("TM1"), "hh:mm:ss") '***2/2/99
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
opmode = " "
savemode = " "
lblmesg = " "
STRSQL = "select * from RISLIB.RISINSTRUC order by insno"
Set rsq = rdb.CreateQuery("qr12", STRSQL)
Set RS = rsq.OpenResultset(2, 3)
Text1.Text = time1
Text2.Text = DATE11

End Sub

Private Sub Form_Unload(cancel As Integer)
Dim RS As rdoResultset
Dim RQ As rdoQuery
rsq.Close
For Each RS In rdb.rdoResultsets
    RS.Close
Next

For Each RQ In rdb.rdoQueries
    RQ.Close
Next
End Sub

Private Sub txtinsno_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 48) Or (KeyAscii > 57) Then
    KeyAscii = 0
  End If
  
End Sub

Private Sub txtinsno_LostFocus()
Dim fndflg As Boolean
 If opmode = "V" Then
  If Len(Trim((txtinsno.Text))) > 0 Then
 
   fndflg = False
   
      RS.MoveFirst
   While Not RS.EOF
     If Trim(RS("insno")) = CInt(Trim(txtinsno)) Then
          fndflg = True
      txtinstext = RS("instext")
      NAVMODE = "V"
     End If
    RS.MoveNext
   Wend
   If fndflg = False Then
        MsgBox "This RECORD is NOT found."
        txtinsno.SetFocus
   End If
   
   
 End If
End If



      
 
End Sub

Private Sub txtinstext_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
End If
i = i + 1
If i = 60 Then
Chr (13)
End If
End Sub
