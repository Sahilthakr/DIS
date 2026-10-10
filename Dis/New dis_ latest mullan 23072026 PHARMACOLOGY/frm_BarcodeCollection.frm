VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_BarcodeCollection 
   Caption         =   "frm_BarcodeCollection"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12030
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1320
      TabIndex        =   8
      Top             =   1200
      Width           =   8415
      Begin VB.TextBox TXTREQNO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3240
         MaxLength       =   11
         TabIndex        =   10
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox cmb_location 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frm_BarcodeCollection.frx":0000
         Left            =   3240
         List            =   "frm_BarcodeCollection.frx":0002
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisation No :-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   960
         TabIndex        =   12
         Top             =   720
         Width           =   2100
      End
      Begin VB.Label lbllocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample Collected At :- "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   2490
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         Height          =   1215
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   8175
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5953
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   16761024
         ForeColorFixed  =   128
         GridColor       =   0
         AllowUserResizing=   3
         FormatString    =   $"frm_BarcodeCollection.frx":0004
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame buttons 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1440
      TabIndex        =   2
      Top             =   7080
      Width           =   8415
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         MouseIcon       =   "frm_BarcodeCollection.frx":009B
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print Report"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmd_back 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Back"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         MouseIcon       =   "frm_BarcodeCollection.frx":03A5
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Back to the Main Screen"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CMD_REFRESH 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Refresh Grid"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         Height          =   855
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   8175
      End
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10440
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   10440
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   1200
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lbl_req 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " SAMPLE COLLECTION CENTRE "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   3240
      TabIndex        =   16
      Top             =   600
      Width           =   4290
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   9840
      TabIndex        =   15
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   9840
      TabIndex        =   14
      Top             =   720
      Width           =   420
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   720
      TabIndex        =   13
      Top             =   0
      Width           =   8595
   End
End
Attribute VB_Name = "frm_BarcodeCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim rs_DHEMCYTO1 As rdoResultset
'Dim qr_DHEMCYTO1 As rdoQuery
'
'Dim rs_RREQUISITN As rdoResultset
'Dim qr_RREQUISITN As rdoQuery
'Dim rs_dhlprosub As rdoResultset
'Dim qr_dhlprosub As rdoQuery
'Dim RS_DMULPROCE As rdoResultset
'Dim qr_dmulproce As rdoQuery
'Dim rs_DBCPROSUB As rdoResultset
'Dim qr_DBCPROSUB As rdoQuery
'Dim rs_SAMPLECODE As rdoResultset
'Dim qr_SAMPLECODE As rdoQuery
'
'Dim ins_dinterface As rdoQuery
'Dim ins_binterface As rdoQuery
'Dim ins_hinterface As rdoQuery
'Dim ins_cinterface As rdoQuery
'Dim ins_spathology As rdoQuery
'Dim ins_DBCINFACE As rdoQuery
'Dim ins_DAUTOPATH As rdoQuery
'Dim ins_AUTOPATHTM As rdoQuery
'
'Dim upd_AUTOPATHTM As rdoQuery
'Dim upd_DREQUISITN As rdoQuery
'Dim upd_RREQUISITN As rdoQuery
'Dim upd_DREQUISITN1 As rdoQuery
'Dim upd_RTIMELOG As rdoQuery
'Dim upd_DAUTOPATH As rdoQuery
'
'Dim rs_rmodality As rdoResultset
'Dim rs_RREQUISITN1 As rdoResultset
'Dim qr_RREQUISITN1 As rdoQuery
'Dim rs_ta101 As rdoResultset
'Dim row, row1 As Integer
'Private Sub cmd_save_Click()
'    Dim Locationsp As String
'    For row = 1 To MSFlexGrid1.Rows
'        If Mid(MSFlexGrid1.TextMatrix(row, 1), 1, 3) = "NZZ" Then
'            Call up_RREQUISITN
'            Call up_DREQUISITN
'            Call up_RTIMELOG
'        Else
'            Call up_RREQUISITN
'            Call up_RTIMELOG
'            Call up_DREQUISITN1
'            If Mid(MSFlexGrid1.TextMatrix(row, 1), 5, 2) = "SP" Then
'                Call spathology
'            End If
'
'            If Mid(MSFlexGrid1.TextMatrix(row, 1), 5, 2) = "HL" Then
'
'            End If
'        End If
'    Next
'End Sub
'Private Sub Form_Load()
'    Set qr_RREQUISITN = rdb.CreateQuery("qr", "select caseno,MODALITYCD from RISLIB.rrequisitn where reqnno=?")
''******* UPDATE
'    STR = "UPDATE DISLIB.AUTOPATHTM SET PATHNO=? WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR=? AND PATHMONTH=? AND PATHDATE=?"
'    Set upd_AUTOPATHTM = rdb.CreateQuery("UPD_A", STR)
'    STR = "UPDATE DISLIB.DAUTOPATH SET PATHNO=? WHERE DEPT_CODE=? AND MODALITYCD=? AND PATHYEAR1=?"
'    Set upd_DAUTOPATH = rdb.CreateQuery("UPD_DA", STR)
'    STR = "UPDATE RISLIB.RTIMELOG SET SAMRECDATE=?,SAMRECTIME=?,CASENO=?,REQNDATE=?, REQNTIME=? WHERE REQNNO=? "
'    Set upd_RTIMELOG = rdb.CreateQuery("UPD_RT", STR)
'    STR = "UPDATE DISLIB.DREQUISITN SET LATERALITS=? WHERE REQNO=?"
'    Set upd_DREQUISITN = rdb.CreateQuery("UPD_DR", STR)
'    STR = "UPDATE DISLIB.DREQUISITN SET PATHNO=? WHERE REQNO=?"
'    Set upd_DREQUISITN1 = rdb.CreateQuery("UPD_DR1", STR)
'    STR = "UPDATE  RISLIB.RREQUISITN SET EXAMDT=?, REQNSTAT=?,EXAMINATN=?,LOCATION=?,APNTDT=? WHERE REQNNO=?"
'    Set upd_RREQUISITN = rdb.CreateQuery("UPD_RR", STR)
''****** INSERT
'    STR = "INSERT INTO INTERFACE.DINTERFACE(CASENO,REQNO,TESTCODE,URGENT,IFLAG,LABNO,FLD2) VALUES(?,?,?,?,?,?,?)"
'    Set ins_dinterface = rdb.CreateQuery("ins_d", STR)
'    STR = "INSERT INTO INTERFACE.BINTERFACE(CASENO,REQNO,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?)"
'    Set ins_binterface = rdb.CreateQuery("ins_b", STR)
'    STR = "INSERT INTO INTERFACE.HINTERFACE(CASENO,REQNO,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?)"
'    Set ins_hinterface = rdb.CreateQuery("ins_h", STR)
'    STR = "INSERT INTO INTERFACE.CINTERFACE(CASENO,REQNO,TESTCODE,URGENT,IFLAG,LABNO) VALUES(?,?,?,?,?,?)"
'    Set ins_cinterface = rdb.CreateQuery("ins_c", STR)
'    STR = "insert into dislib.spathology(path_no,case_no,trans_dt,loct1) values(?,?,?,?)"
'    Set ins_spathology = rdb.CreateQuery("ins_s", STR)
'    STR = "INSERT INTO DISLIB.DAUTOPATH (DEPT_CODE,MODALITYCD,PATHYEAR,PATHYEAR1,PATHNO) VALUES (?,?,?,?,?)"
'    Set ins_DAUTOPATH = rdb.CreateQuery("ins_da", STR)
'    STR = "INSERT INTO DISLIB.AUTOPATHTM (DEPT_CODE,MODALITYCD,PATHYEAR,PATHMONTH,PATHDATE,PATHNO) VALUES (?,?,?,?,?,?)"
'    Set ins_AUTOPATHTM = rdb.CreateQuery("ins_au", STR)
'    STR = "INSERT INTO DISLIB.DBCINFACE(REQNO,LABNO,PROCODE) VALUES(?,?,?)"
'    Set ins_DBCINFACE = rdb.CreateQuery("ins_db", STR)
''****** SELECT
'    Set qr_RREQUISITN = rdb.CreateQuery("qr_rr", "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?")
'    Set qr_dmulproce = rdb.CreateQuery("QR_DM", "SELECT * FROM DISLIB.DMULPROCE WHERE REQNO=?")
'    Set qr_DBCPROSUB = rdb.CreateQuery("qr_db", "SELECT * FROM DISLIB.DBCPROSUB WHERE BCPROCODE=?")
'    Set qr_dhlprosub = rdb.CreateQuery("QR_DH", "select * from dislib.dhlprosub where hlprocode = ?")
'    Set qr_DHEMCYTO1 = rdb.CreateQuery("qr_dhe", "SELECT * FROM DISLIB.DHEMCYTO1 WHERE REQNO = ?")
'
'End Sub
'Private Sub Text1_LostFocus()
'qr_RREQUISITN(0) = req_chglen(Trim(Text1))
'Set rs_RREQUISITN = qr_RREQUISITN.OpenResultset(1, 1)
'Set rs_ta101 = rdb.OpenResultset("select first_name,last_name from patoolslib.ta101 where case_no='" & rs_RREQUISITN(0) & "'", 1, 1)
'Set rs_rmodality = rdb.OpenResultset("select modlitydes from rislib.rmodality where modalitycd='" & rs_RREQUISITN(1) & "'", 1, 1)
'If rs_ta101.RowCount > 0 Then
'    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
'    For row = 1 To MSFlexGrid1.Rows
'        If MSFlexGrid1.TextMatrix(row, 0) = "" Then
'            MSFlexGrid1.TextMatrix(row, 0) = row
'            MSFlexGrid1.TextMatrix(row, 1) = Text1
'            MSFlexGrid1.TextMatrix(row, 2) = rs_RREQUISITN(0)
'            MSFlexGrid1.TextMatrix(row, 3) = Trim(rs_ta101(0)) & " " & Trim(rs_ta101(1))
'            MSFlexGrid1.TextMatrix(row, 4) = rs_rmodality(0)
'            MSFlexGrid1.TextMatrix(row, 5) = "Recevied"
'            MSFlexGrid1.row = row
'            MSFlexGrid1.Col = 6
'            Text2.Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
'            Text2.Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft
'            Text2.Height = MSFlexGrid1.CellHeight
'            Text2.Width = MSFlexGrid1.CellWidth
'            Text2.Visible = True
'            Text2.SetFocus
'            Text2 = ""
'            Exit For
'        End If
'   Next
'Else
'End If
'End Sub
'Private Sub Text2_LostFocus()
'    MSFlexGrid1.TextMatrix(row, 6) = Text2
'    Text2.Visible = False
'    Text1.SetFocus
'End Sub
'Public Sub DINTERFACE()
'    ins_dinterface(0) = MSFlexGrid1.TextMatrix(row, 2)
'    ins_dinterface(1) = req_chglen(Trim(MSFlexGrid1.TextMatrix(row, 1)))
'    ins_dinterface(2) = RS_DMULPROCE!PROCODESUB
'    ins_dinterface(3) = ""
'    ins_dinterface(4) = "1"
'    ins_dinterface(5) = Mid(req_chglen(Trim(MSFlexGrid1.TextMatrix(row, 1))), 4) & rs_DBCPROSUB!SAMPLECD
'    ins_dinterface(6) = Format(Date, "yymmdd") & Trim(MSFlexGrid1.TextMatrix(row, 6))
'    ins_dinterface.Execute
'End Sub
'Public Function req_chglen(str_reqno)
'    Dim res_reqno As String
'    Dim len1 As Integer
'    len1 = Len(Trim(str_reqno))
'    If len1 = 13 Then
'        res_reqno = Left(str_reqno, 3) & "/" & Mid(str_reqno, 4, 2) & "/" & Mid(str_reqno, 6, 2) & "/" & Right(str_reqno, 6)
'    End If
'    If len1 = 16 Then
'        res_reqno = Left(str_reqno, 3) & Mid(str_reqno, 5, 2) & Mid(str_reqno, 8, 2) & Right(str_reqno, 6)
'    End If
'    If len1 = 27 Then
'        res_reqno = Left(str_reqno, 3) & Mid(str_reqno, 5, 2) & Mid(str_reqno, 8, 2) & Mid(str_reqno, 11, 6)
'    End If
'    req_chglen = res_reqno
'End Function
'Public Sub BINTERFACE()
'    ins_binterface(0) = MSFlexGrid1.TextMatrix(row, 2)
'    ins_binterface(1) = req_chglen(MSFlexGrid1.TextMatrix(row, 1))
'    ins_binterface(2) = ""
'    ins_binterface(3) = "1"
'    ins_binterface(4) = Mid(req_chglen(MSFlexGrid1.TextMatrix(row, 1))) & "A"
'    ins_binterface.Execute
'End Sub
'Public Sub HINTERFACE()
'    ins_hinterface(0) = MSFlexGrid1.TextMatrix(row, 2)
'    ins_hinterface(1) = req_chglen(MSFlexGrid1.TextMatrix(row, 1))
'    ins_hinterface(2) = " "
'    ins_hinterface(3) = "1"
'    ins_hinterface(4) = Trim(Mid(req_chglen(MSFlexGrid1.TextMatrix(row, 1)), 4, 10)) + "A"
'    ins_hinterface.Execute
'End Sub
'Public Sub CINTERFACE()
'    ins_cinterface(0) = MSFlexGrid1.TextMatrix(row, 2)
'    ins_cinterface(1) = req_chglen(MSFlexGrid1.TextMatrix(row, 1))
'    ins_cinterface(2) = rs_dhlprosub("procodesub")
'    ins_cinterface(3) = " "
'    ins_cinterface(4) = "1"
'    ins_cinterface(5) = Trim(Mid(hlreqno, 4, 10)) + Trim(rs_dhlprosub("samplecd"))
'    ins_cinterface.Execute
'End Sub
'Public Sub spathology()
'    ins_spathology(0) = UCase(Trim(MSFlexGrid1.TextMatrix(row, 6)))
'    ins_spathology(1) = MSFlexGrid1.TextMatrix(row, 2)
'    ins_spathology(2) = Format(DATE11, "YYYY/MM/DD")
'    ins_spathology(3) = Locationsp
'    ins_spathology.Execute
'End Sub
'Public Sub DAUTOPATH()
'    ins_DAUTOPATH(0) = Mid(MSFlexGrid1.TextMatrix(row, 1), 1, 3)
'    ins_DAUTOPATH(1) = Mid(MSFlexGrid1.TextMatrix(row, 1), 5, 2)
'    ins_DAUTOPATH(2) = Trim(Left(MSFlexGrid1.TextMatrix(row, 2), 2))
'    ins_DAUTOPATH(3) = Mid(Year(Date), 3, 2)
'    ins_DAUTOPATH(4) = 1
'    ins_DAUTOPATH.Execute
'End Sub
'Public Sub AUTOPATHTM()
'    ins_AUTOPATHTM(0) = Mid(MSFlexGrid1.TextMatrix(row, 1), 1, 3)
'    ins_AUTOPATHTM(1) = Mid(MSFlexGrid1.TextMatrix(row, 1), 5, 2)
'    ins_AUTOPATHTM(2) = Trim(Right(DATE11, 2))
'    ins_AUTOPATHTM(3) = Trim(Mid(DATE11, 4, 2))
'    ins_AUTOPATHTM(4) = Trim(Left(DATE11, 2))
'    ins_AUTOPATHTM(5) = 1
'    ins_AUTOPATHTM.Execute
'End Sub
'Public Sub up_AUTOPATHTM()
'    upd_AUTOPATHTM(0) = AUTOPATHNO
'    upd_AUTOPATHTM(1) = Mid(MSFlexGrid1.TextMatrix(row, 1), 1, 3)
'    upd_AUTOPATHTM(2) = Mid(MSFlexGrid1.TextMatrix(row, 1), 5, 2)
'    upd_AUTOPATHTM(3) = Trim(Right(DATE11, 2))
'    upd_AUTOPATHTM(4) = Trim(Mid(DATE11, 4, 2))
'    upd_AUTOPATHTM(5) = Trim(Left(DATE11, 2))
'    upd_AUTOPATHTM.Execute
'End Sub
'Public Sub up_DAUTOPATH()
'    upd_DAUTOPATH(0) = AUTOPATHNO
'    upd_DAUTOPATH(1) = Mid(MSFlexGrid1.TextMatrix(row, 1), 1, 3)
'    upd_DAUTOPATH(2) = Mid(MSFlexGrid1.TextMatrix(row, 1), 5, 2)
'    upd_DAUTOPATH(3) = Trim(Right(DATE11, 2))
'    upd_DAUTOPATH.Execute
'End Sub
'Public Sub up_RTIMELOG()
'    upd_RTIMELOG(0) = Format(txt_date.Text, "YYYY-MM-DD")
'    upd_RTIMELOG(1) = Format(txt_time.Text, "HH:MM:SS")
'    upd_RTIMELOG(2) = msk_caseno.FormattedText
'    upd_RTIMELOG(3) = rs_RREQUISITN("REQNDT")
'    upd_RTIMELOG(4) = rs_RREQUISITN("REQNTM")
'    upd_RTIMELOG(5) = req_chglen(MSFlexGrid1.TextMatrix(row, 1))
'    upd_RTIMELOG.Execute
'End Sub
'Public Sub up_DREQUISITN()
'    upd_DREQUISITN(0) = time1
'    upd_DREQUISITN(1) = req_chglen(MSFlexGrid1.TextMatrix(row, 1))
'    upd_DREQUISITN.Execute
'End Sub
'Public Sub up_RREQUISITN()
'    qr_RREQUISITN(0) = req_chglen(MSFlexGrid1.TextMatrix(row, 1))
'    Set rs_RREQUISITN = qr_RREQUISITN.OpenResultset(1, 1)
'    upd_RREQUISITN(0) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
'    If Trim(rs_RREQUISITN("reqnstat")) = "P" Or Trim(rs_RREQUISITN("reqnstat")) = "C" Or Trim(rs_RREQUISITN("reqnstat")) = "E" Then
'        upd_RREQUISITN(1) = "E"
'    ElseIf Trim(rs_RREQUISITN("reqnstat")) = "T" Then
'        upd_RREQUISITN(1) = "T"
'    ElseIf Trim(rs_RREQUISITN("reqnstat")) = "F" Then
'        upd_RREQUISITN(1) = "F"
'    End If
'    upd_RREQUISITN(2) = "RECEIVED"
'    If Left(Trim(MSFlexGrid1.TextMatrix(row, 2)), 2) = "RF" Or Left(Trim(MSFlexGrid1.TextMatrix(row, 2)), 2) = "RG" Then
'        upd_RREQUISITN(3) = "COLLRCVD"
'        Locationsp = "COLLRCVD"
'    Else
'        If Len(Trim(rs_RREQUISITN("LOCATION"))) = 0 Then  ' made 0 from ""
'            upd_RREQUISITN(3) = "COLLECTED"
'            Locationsp = "COLLECTED"
'        Else
'            upd_RREQUISITN(3) = Trim(rs_RREQUISITN("LOCATION"))
'            If IsNull(Trim(rs_RREQUISITN("LOCATION"))) = True Then
'                Locationsp = ""
'            Else
'                Locationsp = Trim(rs_RREQUISITN("LOCATION"))
'            End If
'        End If
'    End If
'    upd_RREQUISITN(4) = GET_REVDATE(Format(DATE11, "DD/MM/YYYY"))
'    upd_RREQUISITN(5) = req_chglen(MSFlexGrid1.TextMatrix(row, 1))
'    upd_RREQUISITN.Execute
'End Sub
'Public Sub up_DREQUISITN1()
'    If Mid(MSFlexGrid1.TextMatrix(row, 1), 5, 2) = "BC" Or Mid(MSFlexGrid1.TextMatrix(row, 1), 5, 2) = "EL" Or Mid(MSFlexGrid1.TextMatrix(row, 1), 5, 2) = "TM" Then
'        If Len(Trim(MSFlexGrid1.TextMatrix(row, 6))) > 3 Then
'            upd_DREQUISITN1(0) = UCase(Trim(MSFlexGrid1.TextMatrix(row, 6)))
'        ElseIf Len(Trim(MSFlexGrid1.TextMatrix(row, 6))) <= 3 And Len(Trim(MSFlexGrid1.TextMatrix(row, 6))) <> 0 Then
'            upd_DREQUISITN1(0) = Right(DATE11, 2) & Mid(DATE11, 4, 2) & Left(DATE11, 2) & UCase(Trim(MSFlexGrid1.TextMatrix(row, 6)))
'        End If
'    Else
'        upd_DREQUISITN1(0) = UCase(Trim(MSFlexGrid1.TextMatrix(row, 6)))
'    End If
'        upd_DREQUISITN1(1) = req_chglen(MSFlexGrid1.TextMatrix(row, 1))
'        upd_DREQUISITN1.Execute
'End Sub
'Public Sub datetime()
'    Set RS = rdb.OpenResultset("SELECT DISTINCT CURRENT DATE DT1 , CURRENT TIME TM1 FROM PATOOLSLIB.MA107") '***2/2/99
'    DATE11 = RS("dt1")
'    DATE11 = Format(DATE11, "DD/MM/YYYY")
'    time1 = Format(RS("TM1"), "hh:mm:ss")
'    DATETIMEFLAG = True
'End Sub
'
'Private Sub Timer1_Timer()
'    datetime
'    txt_time.Text = time1
'End Sub

Dim SELDEPT_QR As rdoQuery
Dim SELDEPT_RS As rdoResultset
Dim GETCASE_QR As rdoQuery
Dim GETCASE_RS As rdoResultset
Dim GETNAME_QR As rdoQuery
Dim GETNAME_RS As rdoResultset
Dim GETNAMESTAFF_QR As rdoQuery
Dim GETNAMESTAFF_RS As rdoResultset
Dim SELREQNSTAT_QR As rdoQuery
Dim SELREQNSTAT_RS As rdoResultset
Dim UPDRTIME_QR As rdoQuery
Dim UPDRREQ_QR As rdoQuery



Dim REQNOB As String
Dim REQNOBARCODE As String
Dim REQ111 As String
Dim PAT_NAME As String
Dim COLLDATE As String
Dim COLLTIME As String

Dim MODALITYB As String
Dim DEPTB As String
Dim LEN1 As Integer
Dim str As String





Private Sub cmb_location_LostFocus()
If cmb_location.Text = "" Then
    MsgBox "Please Select Location .......!"
    'cmb_location.SetFocus
Else
    TXTREQNO.SetFocus
End If

End Sub

Private Sub cmd_back_Click()

'Me.Hide
If showmain = True Then
   Me.MousePointer = 11
    Me.Hide
    collectionNEW.show
    Me.MousePointer = 0
       
Else
    Me.MousePointer = 11
    Me.Hide
    DIS_MAIN.show
    Me.MousePointer = 0
End If

showmain = False

End Sub

Private Sub cmd_refresh_Click()
MSFlexGrid1.clear
'cmb_location.Text = ""
MSFlexGrid1.Rows = 2
'cmb_location.Text = "PVT. OPD"
'cmb_location.Enabled = False
TXTREQNO.SetFocus
'cmb_location.SetFocus
End Sub

Private Sub Form_Activate()
TXTREQNO.SetFocus
End Sub

Private Sub Form_Load()
str = "SELECT DEPT_CODE,MODLITYDES FROM RISLIB.RMODALITY WHERE MODALITYCD=?"
Set SELDEPT_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set GETCASE_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM PATOOLSLIB.TA101 WHERE CASE_NO=?"
Set GETNAME_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM DISLIB.DREQSTAFFD WHERE REQNO=?"
Set GETNAMESTAFF_QR = rdb.CreateQuery("", str)

str = "SELECT * FROM RISLIB.RREQUISITN WHERE REQNNO=?"
Set SELREQNSTAT_QR = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RTIMELOG SET SAMCOLDATE=?, SAMCOLTIME=? WHERE REQNNO=? "
Set UPDRTIME_QR = rdb.CreateQuery("", str)

str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT=?,LOCATION=?,APNTDT=? WHERE REQNNO=?"
Set UPDRREQ_QR = rdb.CreateQuery("", str)

'TXTREQNO.SetFocus
'cmb_location.AddItem ("GEN.OPD")
cmb_location.AddItem ("PVT.OPD") 'Commented and added by Shraddha on 02/11/09 as told by Mrs.Kalpana
cmb_location.ListIndex = 0
'cmb_location.Enabled = False
'TXTREQNO.SetFocus

datetime
txt_date.Text = DATE11

End Sub

Private Sub Timer1_Timer()
'COMMENTED BY CHARULATA ON 14/09/2010 FOR LESS NO. OF CALLS TO SERVER
'datetime
'txt_time.Text = time1
End Sub

Private Sub TXTREQNO_Change()

'MsgBox "Requisition ..... ", vbOKOnly

If Len(Trim(TXTREQNO.Text)) = 11 Then
    DoEvents
    READDATA
    
End If


End Sub


Public Sub READDATA()

On Error GoTo MYERR

    datetime
    COLLDATE = DATE11
    COLLTIME = time1
    
    
    REQNOB = Trim(TXTREQNO.Text)
    
    REQ111 = Trim(Mid(REQNOB, 3, 8))
    
    MODALITYB = Trim(UCase(Left(REQNOB, 2)))
        
    REQNOBARCODE = MODALITYB & REQ111
        
    SELDEPT_QR(0) = MODALITYB
    Set SELDEPT_RS = SELDEPT_QR.OpenResultset(1, 3)
    DEPTB = Trim(SELDEPT_RS("DEPT_CODE"))
    
      
    ''get reqno
    REQNOB = DEPTB & REQNOBARCODE
    
    SELREQNSTAT_QR(0) = REQNOB
    Set SELREQNSTAT_RS = SELREQNSTAT_QR.OpenResultset(1, 3)
    
    Dim ST1 As String
    ST1 = "C"
    If Trim(SELREQNSTAT_RS("REQNSTAT")) = "C" Then
        DoEvents
        MsgBox "sample already collected....! ", vbOKOnly
        DoEvents
    Else
    
        PAT_NAME = ""
    
        GETCASE_QR(0) = REQNOB
        Set GETCASE_RS = GETCASE_QR.OpenResultset(1, 3)
     
        If Len(Trim(GETCASE_RS("CASENO"))) = 6 Then
        
            GETNAMESTAFF_QR(0) = REQNOB
            Set GETNAMESTAFF_RS = GETNAMESTAFF_QR.OpenResultset(1, 3)

            PAT_NAME = GETNAMESTAFF_RS("PNAME")

        Else
        
            GETNAME_QR(0) = GETCASE_RS("CASENO")
            Set GETNAME_RS = GETNAME_QR.OpenResultset(1, 3)
    
            PAT_NAME = Trim(GETNAME_RS("TAG")) & " " & Trim(GETNAME_RS("FIRST_NAME")) & " " & Trim(GETNAME_RS("MID_NAME")) & " " & Trim(GETNAME_RS("LAST_NAME"))
        End If
    
    
    
            With MSFlexGrid1
            .Rows = .Rows + 1
            .row = .Rows - 2
            
                .TextMatrix(.row, 0) = .row
                .TextMatrix(.row, 1) = REQNOB
                .TextMatrix(.row, 2) = GETCASE_RS("CASENO")
                .TextMatrix(.row, 3) = PAT_NAME
                .TextMatrix(.row, 4) = SELDEPT_RS("MODLITYDES")
                .TextMatrix(.row, 5) = COLLDATE
                .TextMatrix(.row, 6) = COLLTIME
            End With
            rdb.BeginTrans
                
                assign
            rdb.CommitTrans
      End If
            
    DoEvents
    TXTREQNO.Text = ""
    TXTREQNO.SetFocus
Exit Sub

MYERR:
        rdb.RollbackTrans
        MsgBox (Err.Description)
    
End Sub

Public Sub assign()

            UPDRTIME_QR(0) = COLLDATE
            UPDRTIME_QR(1) = COLLTIME
            UPDRTIME_QR(2) = REQNOB
            UPDRTIME_QR.Execute
        
            UPDRREQ_QR(0) = "C"
            UPDRREQ_QR(1) = cmb_location.Text
            UPDRREQ_QR(2) = COLLDATE
            UPDRREQ_QR(3) = REQNOB
            UPDRREQ_QR.Execute
        
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


