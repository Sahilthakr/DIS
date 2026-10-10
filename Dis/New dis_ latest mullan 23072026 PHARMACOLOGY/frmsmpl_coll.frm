VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmsmpl_coll 
   BackColor       =   &H00D8CBD8&
   Caption         =   "Sample Collection "
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form2"
   ScaleHeight     =   9015
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.FileListBox File2 
      Height          =   480
      Left            =   240
      Pattern         =   "*.txt;*.log"
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   360
      Pattern         =   "*.txt;*.log"
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   480
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8BFD8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   720
      TabIndex        =   0
      Top             =   2280
      Width           =   9975
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00D8BFD8&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   960
         TabIndex        =   5
         Top             =   3000
         Width           =   8055
         Begin VB.CommandButton cmd_upload 
            BackColor       =   &H00FFC0C0&
            Caption         =   "&UPLOAD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmd_refresh 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "&REFRESH"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmd_help 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "&HELP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "How to collect sample ?"
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmd_back 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "&BACK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmd_show 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "&COLLECTION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Collect Sample"
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.ComboBox cmb_location 
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
         Left            =   4560
         TabIndex        =   1
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   4095
         Left            =   720
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   8535
      End
      Begin VB.Label lbl_location 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Collected at"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1920
         TabIndex        =   2
         Top             =   1680
         Width           =   2235
      End
   End
   Begin VB.Label lbl_req 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT SAMPLE COLLECTION WITH BARCODE "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   345
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   8190
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
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   1830
      TabIndex        =   3
      Top             =   600
      Width           =   8595
   End
End
Attribute VB_Name = "frmsmpl_coll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Coded by Kanchan to read barcode data from the text file
'and to save and update them in appropriate tables DISLIB.COLLPDT,RISLIB.RTIMELOG,RISLIB.RREQUISITN

Dim SELWARDNAME_QR As rdoQuery
Dim SELWARDNAME_RS As rdoResultset

Dim str As String
Dim WARDNAMEFLAG As Boolean
Dim DEPTFLAG As Boolean
Dim COLLPDTFLAG As Boolean

Dim SELDEPT_QR As rdoQuery
Dim SELDEPT_RS As rdoResultset

Dim INSCOLLPDT_QR As rdoQuery
Dim SELCOLLPDT_QR As rdoQuery
Dim SELCOLLPDT_RS As rdoResultset
Dim UPDCOLLPDT_QR As rdoQuery
Dim UPDRTIME_QR As rdoQuery
Dim UPDRREQ_QR As rdoQuery

Dim fina As String
Dim INSCOLLPDTTXT As rdoQuery
Dim viewtext_qr As rdoQuery
Dim viewtext_rs As rdoResultset
Dim VIEWTXTFLAG As Boolean
Dim UPDCOLLPDTTXT_QR As rdoQuery

Dim WREPORT As Application
Dim DREPORT As Document
Dim WIMPRE As Application
Dim DIMPRE As Document
'************* Shraddha Pandit 25/05/2009 ****************
Dim SEL_REQNSTAT_QR As rdoQuery
Dim SEL_REQNSTAT_RS As rdoResultset
'*********************************************************
Dim tot_smp As Integer 'Shraddha Pandit 26/05/2009

'********** Shraddha Pandit 11/06/2009 *******************
Dim upd_rtime_qr As rdoQuery

Dim sel_rreq_qr As rdoQuery
Dim sel_rreq_rs As rdoResultset

Dim upd_rreq_qr As rdoQuery 'SHRADDHA 09/07/09
'*********************************************************

Private Sub cmb_location_Click()
cmd_Show.Enabled = True
End Sub

Private Sub cmb_location_LostFocus()
cmd_Show.Enabled = True
End Sub

Private Sub cmd_back_Click()
cmb_location.Text = ""
Me.MousePointer = 11
Me.Hide
DIS_MAIN.show
Me.MousePointer = 0

End Sub

Private Sub Cmd_help_Click()
Shell "C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe C:\DISREPORTSEC\HELPDOC.PDF"

End Sub



Private Sub cmd_refresh_Click()
'********* added on 12/06/2009 by Shraddha Pandit ***********
    File2.Pattern = "*.txt;*.log"
    File2.PATH = "C:\disreportsec\data"
    If File2.ListCount <> 0 Then
    For r = 0 To File2.ListCount - 1
        Kill File2.PATH & "\" & File2.List(r)
    Next
    End If

'************************************************************
End Sub

Private Sub cmd_show_Click()

    
   ' MsgBox "Please click on OK button of 'Data Read' to collect the samples."
    
    Shell "C:\disreportsec\data1\Data_Read.exe", vbNormalFocus
    
    'Shell "c:\micro.bat", 0
    
    
    
    MYDELAY (3)
    gettxt
    MYDELAY (3)
    READTXT
        
End Sub
Public Sub gettxt()

Dim fos As Scripting.FileSystemObject
Set fos = New Scripting.FileSystemObject
Dim fol As Folder
Dim fil As File
Dim ti As String
Dim tim As String
Dim STRIP As String
Dim FNAME As String
Dim LEN1 As Integer
    
    STRIP = Winsock1.LocalIP

'Dim fina As String
Set fol = fos.GetFolder("c:\disreportsec\data\")
tim = "0.00"


For Each fil In fol.Files
'fil.DateCreated
ti = fil.DateLastModified
If FormatDateTime(ti, vbShortTime) > FormatDateTime(tim, vbShortTime) Then
    tim = FormatDateTime(ti, vbShortTime)
    fina = fil.name
End If
               
Next

    LEN1 = Len(fina)
    LEN1 = LEN1 - 4
    FNAME = Left(fina, LEN1)

        INSCOLLPDTTXT(0) = FNAME
        INSCOLLPDTTXT(1) = "N"
        INSCOLLPDTTXT(2) = STRIP
        INSCOLLPDTTXT.Execute
        
End Sub

Public Sub READTXT()
Dim req As String
Dim REQNUM As String
Dim coll_dt As String
Dim coll_tm As String
Dim strline As String
Dim arrs() As String
Dim MO As String
Dim STRIP As String
Dim USER_ID As String
Dim finame As String


Dim FNAME As String
Dim LEN1 As Integer
    STRIP = Winsock1.LocalIP
    USER_ID = UID
    FNAME = ""
    
    'Open "C:\Documents and Settings\TMH\Desktop\tata\data\DR.txt" For Input As #1
    
    finame = "C:\disreportsec\data\" & fina
    
    LEN1 = Len(fina)
    LEN1 = LEN1 - 4
    FNAME = Left(fina, LEN1)
           
           
On Error GoTo MYERR:
    
    rdb.BeginTrans
    
    Open finame For Input As #1
    
    While Not EOF(1)
        
        Line Input #1, strline
        arrs = Split(strline, ",")
        req = arrs(0)
        coll_dt = arrs(1)
        coll_tm = arrs(2)
        
        MO = Trim(Left(req, 2))
        req = Trim(Left(req, 10))
                
        SELDEPT_QR(0) = MO
        Set SELDEPT_RS = SELDEPT_QR.OpenResultset(1, 3)
        REQNUM = Trim(SELDEPT_RS("DEPT_CODE")) & Trim(req)
        
        '*************** Shraddha Pandit 25/05/2009 **************
        SEL_REQNSTAT_QR(0) = REQNUM
        Set SEL_REQNSTAT_RS = SEL_REQNSTAT_QR.OpenResultset(1, 3)
        
        If Not SEL_REQNSTAT_RS.EOF Then
        '*********************************************************
            If COLLPDTFLAG = False Then
                SELCOLLPDT_QR(0) = REQNUM
                Set SELCOLLPDT_RS = SELCOLLPDT_QR.OpenResultset(1, 3)
                COLLPDTFLAG = True
            Else
                SELCOLLPDT_QR(0) = REQNUM
                SELCOLLPDT_RS.MoreResults
                SELCOLLPDT_RS.Requery
            End If
         
            If SELCOLLPDT_RS.RowCount > 0 Then
                'UPDCOLLPDT_QR(0) = coll_dt
                'UPDCOLLPDT_QR(1) = coll_tm
                'UPDCOLLPDT_QR(2) = STRIP
                'UPDCOLLPDT_QR(3) = USER_ID
                'UPDCOLLPDT_QR(4) = REQNUM
                'UPDCOLLPDT_QR.Execute
                'MsgBox ("The collection of the sample has already done.")
            Else
                 INSCOLLPDT_QR(0) = REQNUM
                 INSCOLLPDT_QR(1) = cmb_location.Text
                 INSCOLLPDT_QR(2) = coll_dt
                 INSCOLLPDT_QR(3) = coll_tm
                 INSCOLLPDT_QR(4) = STRIP
                 INSCOLLPDT_QR(5) = USER_ID
                     
                 INSCOLLPDT_QR.Execute
             
                 UPDRTIME_QR(0) = coll_dt
                 UPDRTIME_QR(1) = coll_tm
            
                 UPDRTIME_QR(2) = REQNUM
                 UPDRTIME_QR.Execute
             
                 UPDRREQ_QR(0) = "C"
                 UPDRREQ_QR(1) = cmb_location.Text
                 UPDRREQ_QR(2) = coll_dt
                 UPDRREQ_QR(3) = REQNUM
                 UPDRREQ_QR.Execute
                 tot_smp = tot_smp + 1
             '*********** Commented on 26/05/2009 ***************
                 'MsgBox ("The Samples have been Collected....!")
             '***************************************************
             End If
        '*********** 25/05/2009 ***********
        End If
        '**********************************
        
        
        
        '**************** Shraddha 11/06/09 ******************
        sel_rreq_qr(0) = REQNUM
        Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)
        
        If Not sel_rreq_rs.EOF Then
            If sel_rreq_rs("reqnstat") = "P" Then
                upd_rtime_qr(0) = coll_dt
                upd_rtime_qr(1) = coll_tm
                upd_rtime_qr(2) = "C"
                upd_rtime_qr(3) = cmb_location.Text
                upd_rtime_qr(4) = coll_dt
                upd_rtime_qr(5) = REQNUM
                upd_rtime_qr.Execute
            End If
        End If
        
        
        '*****************************************************
        
    Wend
    'MsgBox ("Total " & tot_smp & " Samples have been Collected...!")
    MsgBox ("Samples have been Collected successfully")
    tot_smp = 0
    Close #1
          
    UPDCOLLPDTTXT_QR(0) = FNAME
    UPDCOLLPDTTXT_QR.Execute
    
    rdb.CommitTrans
    
    Exit Sub
    
MYERR:
        rdb.RollbackTrans
        MsgBox Err.Description
End Sub

Private Sub cmd_upload_Click() 'Added on 12/06/09 by Shraddha
datetime
Dim c As Integer
c = 0
Dim sel_barcode_qr As rdoQuery
Dim sel_barcode_rs As rdoResultset
Dim sql_str1 As String

Dim sel_rtime_qr As rdoQuery
Dim sel_rtime_rs As rdoResultset
Dim sql_str2 As String

Dim upd_rtimelog_qr As rdoQuery
Dim sql_str3 As String

sql_str1 = "Select * from dislib.barcode where location='PVT.OPD' and date='" & Format(DATE11, "yyyy-mm-dd") & "'"
Set sel_barcode_qr = rdb.CreateQuery("", sql_str1)

sql_str2 = "Select * from rislib.rtimelog where reqnno=?"
Set sel_rtime_qr = rdb.CreateQuery("", sql_str2)

sql_str3 = "update rislib.rtimelog set samcoldate=?, samcoltime=? where reqnno=?"
Set upd_rtimelog_qr = rdb.CreateQuery("", sql_str3)

'******************************************

Dim sdate, stime As String
    
rdb.BeginTrans
'datetime
Set sel_barcode_rs = sel_barcode_qr.OpenResultset(1, 3)
    
    
While Not sel_barcode_rs.EOF
        
    sel_rtime_qr(0) = sel_barcode_rs("reqno")
    Set sel_rtime_rs = sel_rtime_qr.OpenResultset(1, 3)
    
    If Trim(sel_rtime_rs("samcoldate")) = "" Or IsNull(sel_rtime_rs("samcoldate")) = True Then
    
        sdate = sel_barcode_rs("DATE")
        stime = sel_barcode_rs("time")
        
        upd_rtimelog_qr(0) = sdate
        upd_rtimelog_qr(1) = DateAdd("n", 3, stime)
        upd_rtimelog_qr(2) = sel_barcode_rs("reqno")
        upd_rtimelog_qr.Execute
        c = c + 1
        '**************** update rrequisitn ******************
        sel_rreq_qr(0) = sel_barcode_rs("reqno")
        Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)
    
        If Not sel_rreq_rs.EOF Then
            If Trim(sel_rreq_rs("reqnstat")) = "P" Then
            'Commented and added by Shraddha 09/07/09
'                upd_rtime_qr(0) = sdate
'                upd_rtime_qr(1) = DateAdd(n, 3, stime)
'                upd_rtime_qr(2) = "C"
'                upd_rtime_qr(3) = "PVT.OPD"
'                upd_rtime_qr(4) = sdate
'                upd_rtime_qr(5) = sel_barcode_rs("reqno")
'                upd_rtime_qr.Execute
                upd_rreq_qr(0) = sel_barcode_rs("reqno")
                upd_rreq_qr.Execute
            '******************************************
            End If
        End If
        '*****************************************************
    
    End If
        
    sel_barcode_rs.MoveNext
        
Wend

If c = 0 Then
    MsgBox "successfully uploaded"
'Else
    'MsgBox c & " Records uploaded successfully"
End If

rdb.CommitTrans
    
Exit Sub

End Sub

Private Sub Form_Load()
'********* added on 29/04/2009 by Shraddha Pandit ***********
'If Dir("C:\disreportsec\data") = True Then
    File1.PATH = "C:\disreportsec\data"
    For r = 0 To File1.ListCount - 1
        'MsgBox File1.PATH & "\" & File1.List(r)
        Kill File1.PATH & "\" & File1.List(r)
    Next
'Else
   ' MsgBox "Doesn't Exist"
'End If
'************************************************************

'******************** 25/05/2009 ****************************
'str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO =? AND REQNSTAT='E'"
'Commented and added by Shraddha on 16/09/09
str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO =? AND REQNSTAT='P'"
Set SEL_REQNSTAT_QR = rdb.CreateQuery("", str)
'************************************************************
tot_smp = 0 'Shraddha Pandit 26/05/2009

WARDNAMEFLAG = False

str = "select DISTINCT(LOCATION) from RISLIB.rlocation WHERE (DEPT_CODE <> 'LZZ' OR DEPT_CODE <> 'WAA') ORDER BY LOCATION"
Set SELWARDNAME_QR = rdb.CreateQuery("WARDNAMEselect", str)
 
str = "SELECT DEPT_CODE FROM RISLIB.RMODALITY WHERE MODALITYCD=?"
Set SELDEPT_QR = rdb.CreateQuery("DEPARTMENTNAME", str)

str = "INSERT INTO DISLIB.COLLPDT (REQNO,WARD_NAME,COLL_DATE,COLL_TIME,IP_ADD,USER_ID) VALUES (?,?,?,?,?,?)"
Set INSCOLLPDT_QR = rdb.CreateQuery("INSERTCOLLPDT", str)

str = "UPDATE RISLIB.RTIMELOG SET SAMCOLDATE=?, SAMCOLTIME=? WHERE REQNNO=? "
'STR = "UPDATE RISLIB.RTIMELOG SET SAMCOLDATE=?, SAMCOLTIME=? WHERE REQNNO=?"
Set UPDRTIME_QR = rdb.CreateQuery("UPDATERTIMELOG", str)

str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT=?,LOCATION=?,APNTDT=? WHERE REQNNO=?"
Set UPDRREQ_QR = rdb.CreateQuery("UPDATERREQUISITN", str)
 
'************************* 11/06/09 Shraddha *******************

str = "UPDATE RISLIB.RTIMELOG SET SAMCOLDATE=?, SAMCOLTIME=?,REQNSTAT=?,LOCATION=?,APNTDT=? WHERE REQNNO=? "
Set upd_rtime_qr = rdb.CreateQuery("", str)

str = "Select * from rislib.rrequisitn where reqnno=?"
Set sel_rreq_qr = rdb.CreateQuery("", str)
'*************** Shraddha 09/07/09 ******************
str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT='C' where REQNNO=?"
Set upd_rreq_qr = rdb.CreateQuery("", str)
'*******************************************************
 
str = "SELECT * FROM DISLIB.COLLPDT WHERE REQNO=?"
Set SELCOLLPDT_QR = rdb.CreateQuery("SELECTCOLLPDT", str)

str = "UPDATE DISLIB.COLLPDT SET COLL_DATE=?,COLL_TIME=?,IP_ADD=?,USER_ID=? WHERE REQNO=?"
Set UPDCOLLPDT_QR = rdb.CreateQuery("UPDATECOLLPDT", str)

str = "INSERT INTO DISLIB.COLLPDTTXT (FILENAME,FLAG,IP) VALUES (?,?,?)"
Set INSCOLLPDTTXT = rdb.CreateQuery("INSERTCOLLPDTTXT", str)

str = "SELECT * FROM DISLIB.COLLPDTTXT WHERE FILENAME=?"
Set viewtext_qr = rdb.CreateQuery("VIEWTEXTQRY", str)

str = "UPDATE DISLIB.COLLPDTTXT SET FLAG='Y' WHERE FILENAME=?"
Set UPDCOLLPDTTXT_QR = rdb.CreateQuery("UPDATECOLLPDTTXT", str)

 If WARDNAMEFLAG = False Then
    Set SELWARDNAME_RS = SELWARDNAME_QR.OpenResultset(1, 3)
    WARDNAMEFLAG = True
 Else
    SELWARDNAME_RS.MoreResults
    SELWARDNAME_RS.Requery
 End If
 
    If SELWARDNAME_RS.RowCount > 0 Then
    
        While Not SELWARDNAME_RS.EOF
     
            cmb_location.AddItem Trim(SELWARDNAME_RS("LOCATION"))
                  SELWARDNAME_RS.MoveNext
        Wend
        
        cmb_location.Text = "PVT.OPD" '**** shweta raut on 14-may-09 as told by kalpana madam
        
    End If
    
DELFILE
cmb_location.Enabled = True
cmd_back.Enabled = True
cmd_Show.Enabled = True 'False '**** shweta raut on 14-may-09 as told by kalpana madam
End Sub
Public Sub DELFILE()
Dim fos1 As Scripting.FileSystemObject
Set fos1 = New Scripting.FileSystemObject
Dim fol1 As Folder
Dim fil1 As File
Dim ti1 As String
Dim tim1 As String
Dim CHECKFLAG As String
Dim FNAME As String
Dim LEN1 As Integer



'Dim fina As String
Set fol1 = fos1.GetFolder("c:\disreportsec\data\")
tim1 = "0.00"



For Each fil1 In fol1.Files
'fil.DateCreated
ti1 = fil1.DateLastModified
Dim ST1 As String
Dim st2 As String

If Left(FormatDateTime(ti1, vbGeneralDate), 2) <> Left(Date, 2) Then

    ST = FormatDateTime(ti1, vbShortDate)
    fina = fil1.name
    FNAME = ""
    LEN1 = Len(fina)
    LEN1 = LEN1 - 4
    FNAME = Left(fina, LEN1)
             
                
        If VIEWTXTFLAG = False Then
            viewtext_qr(0) = FNAME
            Set viewtext_rs = viewtext_qr.OpenResultset(1, 3)
            VIEWTXTFLAG = True
        Else
            viewtext_qr(0) = FNAME
            viewtext_rs.MoreResults
            viewtext_rs.Requery
        End If
    
    If viewtext_rs.RowCount > 0 Then
    
        CHECKFLAG = viewtext_rs("FLAG")
    Else
        CHECKFLAG = "Y"
    End If
    
    If CHECKFLAG = "Y" Then
        fil1.delete
    ElseIf CHECKFLAG = "N" Then
        
        READTXT
        fil1.delete
    End If
End If

Next

End Sub
Public Sub MYDELAY(inttime As String)



Dim sttime
Dim endtime1 As Long
Dim TESTTIME As Long


'sttime = Timer

endtime1 = Timer + inttime
TESTTIME = Timer

Do While Timer < endtime1
 '   Text1 = endtime1
 '   Text2 = TESTTIME
    DoEvents
  '  TESTTIME = Timer
Loop


End Sub

