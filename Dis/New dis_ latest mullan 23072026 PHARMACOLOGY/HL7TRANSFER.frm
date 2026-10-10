VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form HL7TRANSFER 
   Caption         =   "HL7TRANSFER"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12345
   LinkTopic       =   "Form2"
   ScaleHeight     =   9840
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "SEND NOW"
      Height          =   495
      Left            =   7920
      TabIndex        =   24
      Top             =   7200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SEND ORDER"
      Height          =   615
      Left            =   3120
      TabIndex        =   23
      Top             =   7920
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock win_client 
      Left            =   11760
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
      Caption         =   "BACK"
      Height          =   615
      Left            =   9960
      TabIndex        =   22
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "CANCEL"
      Height          =   615
      Left            =   7800
      TabIndex        =   21
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "STATUS"
      Height          =   615
      Left            =   5880
      TabIndex        =   4
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PACS REPORTS"
      Height          =   495
      Left            =   -360
      TabIndex        =   3
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PACS ORDERS"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "PACS ORDERS"
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10935
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   1800
         TabIndex        =   18
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   2280
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   1680
         Width           =   5775
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   5520
         TabIndex        =   12
         Top             =   960
         Width           =   2415
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "??>/#####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   6120
         TabIndex        =   8
         Top             =   360
         Width           =   3735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "PATH NO."
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "PROCEDURE NAME"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "PATIENT NAME"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "REQUISITION NO."
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "CASE NO"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "MODALITY"
         Height          =   255
         Left            =   5160
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "DEPTARTMENT"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "REPORTS AND  ADDENDUM"
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   11055
      Begin VB.CommandButton Command5 
         Caption         =   "PRINT"
         Height          =   495
         Left            =   3600
         TabIndex        =   20
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2895
         Left            =   0
         TabIndex        =   19
         Top             =   480
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   5106
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         FormatString    =   "S NO  |^   CASE NO  |^    PATH NO.   |^ REQUISITION NO. |^PROVISIONAL REPORT   |^   FINAL REPORT    |^     ADDENDUM   "
      End
   End
End
Attribute VB_Name = "HL7TRANSFER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim sel_deptqr As rdoQuery
Dim sel_moadlityqr As rdoQuery
Dim sel_deptrs As rdoResultset
Dim sel_modalityrs As rdoResultset
Dim req_qr  As rdoQuery
Dim req_rs As rdoResultset
Dim sel_logqr As rdoQuery
Dim sel_logrs As rdoResultset
Dim sel_ta101qr As rdoQuery
Dim sel_ta101rs As rdoResultset
Dim sel_ma101qr As rdoQuery
Dim sel_ma101rs As rdoResultset
Dim reserrq As rdoQuery
Dim reserrs As rdoResultset
Dim sel_disreqqr As rdoQuery
Dim sel_disreqrs As rdoResultset
Dim insas400_hl7_rq As rdoQuery
Dim insas400_hl7_rs As rdoResultset
Dim insertlog As rdoQuery
Dim HL7DISVWREP_QR As rdoQuery
Dim HL7DISVWREP_RS As rdoResultset
Dim HL7DISVWADD_QR As rdoQuery
Dim HL7DISVWADD_RS As rdoResultset
Dim mmode As String




Private Sub Combo1_Change()
'Dim dept As Integer
'Dim depcode As String
'    dept = InStr(Combo1.Text, "$")
'    DEPTCODE = Mid(Combo1.Text, dept + 1, 3)
'sel_moadlityqr(0) = DEPTCODE
'Set sel_modalityrs = sel_moadlityqr.OpenResultset(1, 1)
'sel_modalityrs.MoveFirst
'Combo2.clear
'If sel_modalityrs.RowCount > 0 Then
'  Do While sel_modalityrs.EOF = False
'   Combo2.AddItem Trim(sel_modalityrs("modalityname")) + "$" + Trim(sel_modalityrs("modalitycd"))
'    sel_modalityrs.MoveNext
'  Loop
'
'End If

End Sub

Private Sub Combo1_Click()
Dim dept As Integer
Dim depcode As String
    dept = InStr(Combo1.Text, "$")
    DEPTCODE = Mid(Combo1.Text, dept + 1, 3)
sel_moadlityqr(0) = DEPTCODE
Set sel_modalityrs = sel_moadlityqr.OpenResultset(1, 1)
If sel_modalityrs.RowCount > 0 Then

sel_modalityrs.MoveFirst
End If


Combo2.clear
If sel_modalityrs.RowCount > 0 Then
  Do While sel_modalityrs.EOF = False
  If Trim(sel_modalityrs("modalitycd")) = "SP" Then
   
   Combo2.AddItem Trim(sel_modalityrs("modlitydes")) + "$" + Trim(sel_modalityrs("modalitycd"))
   Combo2.Text = Trim(sel_modalityrs("modlitydes")) + "$" + Trim(sel_modalityrs("modalitycd"))
  End If
  
    sel_modalityrs.MoveNext
  Loop
  
End If
End Sub

Private Sub Combo3_Click()
sel_ta101qr(0) = MaskEdBox1.FormattedText
Set sel_ta101rs = sel_ta101qr.OpenResultset(1, 1)
 If sel_ta101rs.RowCount > 0 Then
  Text1.Text = Trim(sel_ta101rs("tag")) + " " + Trim(sel_ta101rs("first_name")) + " " + Trim(sel_ta101rs("mid_name")) + " " + Trim(sel_ta101rs("last_name"))
 End If
 reserrq(0) = req_chglen(Combo3.Text)
 Set reserrs = reserrq.OpenResultset(1, 1)
If reserrs.RowCount > 0 Then

 
 sel_ma101qr(0) = reserrs("ser_code")
 Set sel_ma101rs = sel_ma101qr.OpenResultset(1, 1)
 If sel_ma101rs.RowCount > 0 Then
   Text2.Text = Trim(sel_ma101rs("ser_det")) + "$" + Trim(sel_ma101rs("ser_code"))
   
 End If
End If
sel_disreqqr(0) = req_chglen(Combo3.Text)
Set sel_disreqrs = sel_disreqqr.OpenResultset(1, 1)
If sel_disreqrs.RowCount > 0 Then
   Text3.Text = Trim(sel_disreqrs("pathno"))
End If

End Sub

Private Sub Command1_Click()
mmode = "ORD"
Frame2.Enabled = True
Frame1.Enabled = False

Set sel_deptrs = sel_deptqr.OpenResultset(1, 1)
Combo1.clear
sel_deptrs.MoveFirst
Do While sel_deptrs.EOF = False
If Trim(sel_deptrs("dept_code")) = "FZZ" Then

 Combo1.AddItem Trim(sel_deptrs("dept_name")) + "$" + Trim(sel_deptrs("dept_code"))
 Combo1.Text = Trim(sel_deptrs("dept_name")) + "$" + Trim(sel_deptrs("dept_code"))
End If

sel_deptrs.MoveNext
Loop
Combo2.clear
Combo3.clear

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
MaskEdBox1.Text = ""
MSFlexGrid1.clear
MSFlexGrid1.FormatString = "S NO  |^   CASE NO  |^    PATH NO.   |^ REQUISITION NO. |^PROVISIONAL REPORT   |^   FINAL REPORT    |^     ADDENDUM   "

End Sub

Private Sub Command2_Click()
Frame2.Enabled = False
Frame1.Enabled = True
mmode = "REP"
Command6.Caption = "SEND REPORT/S"
  MSFlexGrid1.clear
  MSFlexGrid1.FormatString = "S NO  |^   CASE NO  |^    PATH NO.   |^ REQUISITION NO. |^ NO. OF DAYS |^  REPORT PRESENT    |^ SENT TO PACS |^              AUTHOR"
  
  MSFlexGrid1.row = 1
Set HL7DISVWREP_RS = HL7DISVWREP_QR.OpenResultset(1, 1)
If HL7DISVWREP_RS.RowCount > 0 Then
  MSFlexGrid1.clear
  MSFlexGrid1.FormatString = "S NO  |^   CASE NO  |^    PATH NO.   |^ REQUISITION NO. |^ NO. OF DAYS |^  REPORT PRESENT    |^ SENT TO PACS |^              AUTHOR"
  
  MSFlexGrid1.row = 1
  
 MSFlexGrid1.Rows = HL7DISVWREP_RS.RowCount + 1
   HL7DISVWREP_RS.MoveFirst
  Do While HL7DISVWREP_RS.EOF = False
   For i = 1 To HL7DISVWREP_RS.RowCount
      MSFlexGrid1.row = i
       MSFlexGrid1.TextMatrix(i, 0) = CStr(i)
       MSFlexGrid1.TextMatrix(i, 1) = HL7DISVWREP_RS("caseno")
       MSFlexGrid1.TextMatrix(i, 2) = HL7DISVWREP_RS("path_no")
       MSFlexGrid1.TextMatrix(i, 3) = HL7DISVWREP_RS("reqno")
       MSFlexGrid1.TextMatrix(i, 4) = HL7DISVWREP_RS("noday")
       MSFlexGrid1.TextMatrix(i, 5) = HL7DISVWREP_RS("ordsend")
       MSFlexGrid1.TextMatrix(i, 6) = "N"  'HL7DISVWREP_RS("ordsend")
       MSFlexGrid1.TextMatrix(i, 7) = "  " + HL7DISVWREP_RS("freportby") + " "
   Next
    HL7DISVWREP_RS.MoveNext
       
  Loop
End If
 
End Sub

Private Sub Command3_Click()
Dim MSTR1 As String

Frame2.Enabled = False
Frame1.Enabled = True
Frame1.Visible = True
MSFlexGrid1.Visible = True

mmode = "ADD"
Command6.Caption = "SEND ADDENDUM/S"
MSFlexGrid1.clear
  MSFlexGrid1.FormatString = "S NO  |^   CASE NO  |^    PATH NO.   |^ REQUISITION NO. |^ NO. OF DAYS |^  ADDENDUM  PRESENT    |^ SENT TO PACS |^              AUTHOR"
  MSFlexGrid1.FormatString = "S NO  |^   CASE NO  |^    PATH NO.   |^ REQUISITION NO. |^PROVISIONAL REPORT   |^   FINAL REPORT    |^     ADDENDUM   "
  MSFlexGrid1.row = 1
Set HL7DISVWADD_RS = HL7DISVWADD_QR.OpenResultset(1, 1)
If HL7DISVWADD_RS.RowCount > 0 Then
  MSFlexGrid1.clear
  MSFlexGrid1.FormatString = "S NO  |^   CASE NO  |^    PATH NO.   |^ REQUISITION NO. |^PROVISIONAL REPORT   |^   FINAL REPORT    |^     ADDENDUM   "
  
  MSFlexGrid1.row = 1
  
 MSFlexGrid1.Rows = HL7DISVWADD_RS.RowCount + 1
   HL7DISVWADD_RS.MoveFirst
  Do While HL7DISVWADD_RS.EOF = False
   For i = 1 To HL7DISVWADD_RS.RowCount
      MSFlexGrid1.row = i
       MSFlexGrid1.TextMatrix(i, 0) = CStr(i)
       MSFlexGrid1.TextMatrix(i, 1) = HL7DISVWADD_RS("caseno")
       MSFlexGrid1.TextMatrix(i, 2) = HL7DISVWADD_RS("path_no")
       MSFlexGrid1.TextMatrix(i, 3) = HL7DISVWADD_RS("reqno")
       If (IsNull(HL7DISVWADD_RS("R_DATE")) = True) Then
         MSTR1 = ""
       Else
            MSTR1 = CStr(HL7DISVWADD_RS("R_DATE"))
            MSFlexGrid1.TextMatrix(i, 4) = MSTR1
       End If
       
            
       
       If (IsNull(HL7DISVWADD_RS("FR_DATE")) = True) Then
       
          MSTR1 = ""
       Else
            MSTR1 = CStr(HL7DISVWADD_RS("FR_DATE"))
            MSFlexGrid1.TextMatrix(i, 5) = MSTR1
       End If
        'HL7DISVWADD_RS("noday")
       If (IsNull(HL7DISVWADD_RS("ADD_DATE")) = True) Then
          MSTR1 = ""
       Else
            MSTR1 = CStr(HL7DISVWADD_RS("ADD_DATE"))
            MSFlexGrid1.TextMatrix(i, 6) = MSTR1
       End If
       'MSFlexGrid1.TextMatrix(i, 5) = HL7DISVWADD_RS("ordsend")
       
       'MSFlexGrid1.TextMatrix(i, 6) = "N"  'HL7DISVWADD_RS("ordsend")
      ' MSFlexGrid1.TextMatrix(i, 7) = HL7DISVWADD_RS("freportby")
   
    HL7DISVWADD_RS.MoveNext
   Next
  Loop
End If
 
End Sub

Private Sub Command4_Click()
On Error GoTo errmsg

Dim caseno_hl7 As String
    Dim reqnno_hl7 As String
'    caseno_hl7 = msk_caseno.FormattedText
'    reqnno_hl7 = cbo_reqno.Text

    Dim msgtype_hl7 As String

    Dim str_hl7 As String
    Dim uniqueid_hl7 As String
      
If Len(Trim(Combo1.Text)) = 0 Then
     MsgBox "Department is Not Seleected"
     Exit Sub
End If
If Len(Trim(Combo2.Text)) = 0 Then
     MsgBox "Modality is Not Seleected"
     Exit Sub
End If

If Len(Trim(Combo3.Text)) = 0 Then
     MsgBox "Requisition is Not Seleected"
     Exit Sub
End If
If Len(Trim(Text3.Text)) = 0 Then
     MsgBox "Path Number is not Present"
     Exit Sub
End If

  '  Dim insas400_hl7_rq As rdoQuery
  '  Dim insas400_hl7_rs As rdoResultset
'    Dim caseno_hl7 As String
'    Dim reqnno_hl7 As String
''    caseno_hl7 = msk_caseno.FormattedText
''    reqnno_hl7 = cbo_reqno.Text
'
'    Dim msgtype_hl7 As String
'
'    Dim str_hl7 As String
'    Dim uniqueid_hl7 As String
'
'     cmd_Send_ord.Enabled = False
'     cmd_sendreport.Enabled = True
'     cmd_addendum.Enabled = True
    '*****Insering into HL7 table **24122003
    caseno_hl7 = MaskEdBox1.FormattedText
    reqnno_hl7 = Combo3.Text
    msgtype_hl7 = "Neworder" 'Form1.lbl_type.Tag
    str_hl7 = WIN_CLIENT.LocalIP
'    selectlog(0) = req_chglen(reqnno_hl7)
'        If selectlogflag Then
'            rslog.MoreResults
'            rslog.Requery
'        Else
'           Set rslog = selectlog.OpenResultset(1, 1)
'           selectlogflag = True
'        End If
'        If Not rslog.EOF Then
'            MsgBox "Yor have already Send an order to PACS on date " & Trim(rslog("OD_DATE")) & " at time " & Trim(rslog("OD_TIME")) & " ", vbInformation
'            Exit Sub
'        End If
    Set DTTIM = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE  , CURRENT TIME  FROM PATOOLSLIB.MA107") '***2/2/99
        uniqueid_hl7 = str_hl7 & "-" & Format(DTTIM(0), "dd/mm/yyyy") & "-" & Format(DTTIM(1), "hh:mm:ss")
       ' Set insas400_hl7_rq = rdb.CreateQuery("ins as4001_r", "insert into hl7lib.hl7tab3(UNQMSGID,MSGTYPID,CASENO,CLNTIP,CLNTDATE,CLNTTIME,REQNNO,PROCBRDGYN) values(?,?,?,?,?,?,?,?)")
rdb.BeginTrans
        insas400_hl7_rq(0) = uniqueid_hl7
        insas400_hl7_rq(1) = msgtype_hl7
        insas400_hl7_rq(2) = MaskEdBox1.FormattedText
        insas400_hl7_rq(3) = WIN_CLIENT.LocalIP
        insas400_hl7_rq(4) = Format(DTTIM(0), "dd/mm/yyyy")
        insas400_hl7_rq(5) = Format(DTTIM(1), "hh:mm:ss")
        insas400_hl7_rq(6) = req_chglen(reqnno_hl7)
        insas400_hl7_rq(7) = "N"
        insas400_hl7_rq.Execute
         delay (2)
         msgtype_hl7 = "Exampull" '
         Set DTTIM = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE  , CURRENT TIME  FROM PATOOLSLIB.MA107") '***2/2/99
        uniqueid_hl7 = str_hl7 & "-" & Format(DTTIM(0), "dd/mm/yyyy") & "-" & Format(DTTIM(1), "hh:mm:ss")
        insas400_hl7_rq(0) = uniqueid_hl7
        insas400_hl7_rq(1) = msgtype_hl7
        insas400_hl7_rq(2) = MaskEdBox1.FormattedText
        insas400_hl7_rq(3) = WIN_CLIENT.LocalIP
        insas400_hl7_rq(4) = Format(DTTIM(0), "dd/mm/yyyy")
        insas400_hl7_rq(5) = Format(DTTIM(1), "hh:mm:ss")
        insas400_hl7_rq(6) = req_chglen(reqnno_hl7)
        insas400_hl7_rq(7) = "N"
        insas400_hl7_rq.Execute
        'FOR INSERTING INTO LOG FILE
      '  Set insertlog = rdb.CreateQuery("insertlog ", "insert into DISLIB.PACSLOG(CASENO,REQNO,ORDSEND,OD_DATE, OD_TIME) values(?,?,?,?,?)")
          insertlog(0) = MaskEdBox1.FormattedText
          insertlog(1) = req_chglen(reqnno_hl7)
          insertlog(2) = "Y"
          insertlog(3) = Format(DTTIM(0), "dd/mm/yyyy")
          insertlog(4) = Format(DTTIM(1), "hh:mm:ss")
          insertlog(5) = "N"
          insertlog(6) = Null
          insertlog(7) = Null
          insertlog(8) = "N"
          insertlog(9) = Null
          insertlog(10) = Null
          insertlog(11) = Trim(Text3.Text)
          insertlog(12) = Mid(req_chglen(Combo3.Text), 1, 3)
          insertlog(13) = Mid(req_chglen(Combo3.Text), 4, 2)
          
          insertlog.Execute
rdb.CommitTrans
        MsgBox "NewOrder successfully sent to AS/400 - HL7 Table", vbInformation
        Exit Sub
        
errmsg:
  MsgBox "ERROR WHILE SENDING ORDER"
  MsgBox Err.Description
  rdb.RollbackTrans
  

End Sub

Private Sub Command8_Click()
Combo1.clear
Combo2.clear
Combo3.clear

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
MaskEdBox1.Text = ""
MSFlexGrid1.clear
MSFlexGrid1.FormatString = "S NO  |^   CASE NO  |^    PATH NO.   |^ REQUISITION NO. |^PROVISIONAL REPORT   |^   FINAL REPORT    |^     ADDENDUM   "

End Sub

Private Sub Command9_Click()
'End
'Me.Hide
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Me.Hide
DIS_MAIN.show

'Unload Me

End Sub

Private Sub Form_Load()
Frame1.Enabled = False
Frame2.Enabled = False

str = "select dept_name, dept_code from PATOOLSLIB.ma106 WHERE (DEPT_CODE='FZZ' OR DEPT_CODE='FEC' OR DEPT_CODE='GZZ' OR DEPT_CODE='CAZ' OR DEPT_CODE='CBZ' OR DEPT_CODE='UZZ' OR DEPT_CODE='NZZ' OR DEPT_CODE='LZZ') "
Set sel_deptqr = rdb.CreateQuery("sel1_deptqr111", str)
str = "select modlitydes,modalitycd from rislib.rmodality WHERE DEPT_CODE=? "
Set sel_moadlityqr = rdb.CreateQuery("sel1_moadlityqr111", str)
Set req_qr = rdb.CreateQuery("1req1", "select * from Rislib.RREQUISITN where caseno=?  AND MODALITYCD=? ORDER BY REQNDT desc ")
str = "select * from dislib.pacslog where reqno = ?"
Set sel_logqr = rdb.CreateQuery("logqr", str)
str = "select * from patoolslib.ta101 where case_no = ?"
'Set sel_ta101qr = rdb.CreateQuery("selta101", STR)
Set sel_ta101qr = rdb.CreateQuery("SEL1", str)
str = "select * from patoolslib.ma101 where ser_code = ?"
Set sel_ma101qr = rdb.CreateQuery("selma101", str)
str = "select ser_code from rislib.rrequisitn where reqnno = ?"
Set reserrq = rdb.CreateQuery("resser", str)
str = "select * from dislib.drequisitn where reqno = ?"
Set sel_disreqqr = rdb.CreateQuery("sel_disreq", str)
Set insas400_hl7_rq = rdb.CreateQuery("ins as4001_r", "insert into hl7lib.hl7tab3(UNQMSGID,MSGTYPID,CASENO,CLNTIP,CLNTDATE,CLNTTIME,REQNNO,PROCBRDGYN) values(?,?,?,?,?,?,?,?)")
Set insertlog = rdb.CreateQuery("insertlog ", "insert into DISLIB.PACSLOG(CASENO,REQNO,ORDSEND,OD_DATE, OD_TIME,repsend,r_date,r_time,addendum,add_date,add_time,path_no,deptcode,modcd) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?)")
str = "SELECT * FROM DISLIB.HL7DISVW WHERE ORDSEND = 'Y' AND REPSEND = 'N'"
Set HL7DISVWREP_QR = rdb.CreateQuery("HL7ERP1", str)
'STR = "SELECT * FROM DISLIB.HL7DISVW WHERE ORDSEND = 'Y' AND REPSEND = 'Y' AND ADDENDUM = 'N'"
str = "SELECT * FROM DISLIB.PACSLOG"
Set HL7DISVWADD_QR = rdb.CreateQuery("HL7ADD11P", str)


End Sub

Private Sub MaskEdBox1_LostFocus()
Dim CASENO As String
Dim modpos As Integer
Dim modi As String

'On Error GoTo MSG:
If Len(Trim(MaskEdBox1.Text)) = "" Then
Exit Sub
End If

    CASENO = MaskEdBox1.Text
    MaskEdBox1.Text = UCase(padzero(5, "C", CASENO, 3))

   modpos = InStr(Combo2.Text, "$")
    modi = Mid(Combo2.Text, modpos + 1, 3)
req_qr(0) = MaskEdBox1.FormattedText
req_qr(1) = modi
Set req_rs = req_qr.OpenResultset(1, 1)
Combo3.clear

If req_rs.RowCount > 0 Then
 req_rs.MoveFirst
 Do While req_rs.EOF = False
      sel_logqr(0) = req_rs("reqnno") 'Combo3.Text
      Set sel_logrs = sel_logqr.OpenResultset(1, 1)
      Do While sel_logrs.StillExecuting
         DoEvents
      Loop
      If sel_logrs.RowCount = 0 Then
        Combo3.AddItem req_chglen(req_rs("reqnno"))
      End If
     req_rs.MoveNext
 Loop
'' patient name


End If
End Sub


Public Sub delay(inttime As String)
Dim sttime
Dim endtime
sttime = Timer
endtime = sttime + inttime

Do While Timer < endtime
 DoEvents
Loop


End Sub
