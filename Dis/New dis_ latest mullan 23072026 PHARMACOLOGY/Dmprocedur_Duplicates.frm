VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form DMPRO 
   Caption         =   "Form2"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form2"
   ScaleHeight     =   9075
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton cmd_delete 
      Caption         =   "Delete Duplicates"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   8400
      Width           =   2655
   End
   Begin VB.CommandButton cmd_Show 
      Caption         =   "Show Duplicate records"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   8400
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7575
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   13361
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "31/08/2010"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   9600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "DMPROCEDUR"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "DMPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim str As String
'Dim sel_dup_dmpro_qr As rdoQuery
'Dim sel_dup_dmpro_rs As rdoResultset
'
'Dim sel_dmpro_qr As rdoQuery
'Dim sel_dmpro_rs As rdoResultset
'
'Dim del_dmpro_qr As rdoQuery
'
'Dim INS_DMRRO_qr As rdoQuery
'
'Dim reqno, PRO_CODE, OBSERVALUE, OBSER_TYPE, Source, ENTRYBY, REMARKS, UOM, NRANGE, PVALUE As String
'
'Dim I As Long
'
'Private Sub cmd_exit_Click()
'End
'End Sub
'
'Private Sub Form_Load()
'str = "SELECT REQNO, PRO_CODE, COUNT(*) FROM DISLIB.DMPROCEDUR GROUP BY REQNO,PRO_CODE Having count(reqno) > 1"
'Set sel_dup_dmpro_qr = rdb.CreateQuery("", str)
'
''REQNO PRO_CODE  OBSERVALUE  OBSER_TYPE  SOURCE  ENTRYBY REMARKS UOM NRANGE PVALUE
'
'str = "SELECT * FROM DISLIB.DMPROCEDUR where reqno=? AND PRO_CODE=?"
'Set sel_dmpro_qr = rdb.CreateQuery("", str)
'
'str = "delete from DISLIB.DMPROCEDUR where reqno=? AND PRO_CODE=?"
'Set del_dmpro_qr = rdb.CreateQuery("", str)
'
'str = "INSERT INTO DISLIB.DMPROCEDUR VALUES(?,?,?,?,?,?,?,?,?,?)"
'Set INS_DMRRO_qr = rdb.CreateQuery("", str)
'End Sub
'
'Private Sub cmd_show_Click()
'Set sel_dup_dmpro_rs = sel_dup_dmpro_qr.OpenResultset(1, 3)
'MSFlexGrid1.Rows = sel_dup_dmpro_rs.RowCount + 1
'MSFlexGrid1.Cols = 4
'
'MSFlexGrid1.TextMatrix(0, 0) = "REQNO"
'MSFlexGrid1.TextMatrix(0, 1) = "Pro_code"
'MSFlexGrid1.TextMatrix(0, 2) = "Count"
'
'MSFlexGrid1.ColWidth(0) = 500
'MSFlexGrid1.ColWidth(1) = 2000
'MSFlexGrid1.ColWidth(2) = 1000
'MSFlexGrid1.ColWidth(3) = 1000
'
'For I = 1 To sel_dup_dmpro_rs.RowCount
'    MSFlexGrid1.TextMatrix(I, 0) = I
'    MSFlexGrid1.TextMatrix(I, 1) = sel_dup_dmpro_rs(0)
'    MSFlexGrid1.TextMatrix(I, 2) = sel_dup_dmpro_rs(1)
'    MSFlexGrid1.TextMatrix(I, 3) = sel_dup_dmpro_rs(2)
'   sel_dup_dmpro_rs.MoveNext
'Next I
'
'End Sub
'
'Private Sub cmd_delete_Click()
'sel_dup_dmpro_rs.MoveFirst
'For I = 1 To sel_dup_dmpro_rs.RowCount
''While Not sel_dup_dmpro_rs.EOF
'    sel_dmpro_qr(0) = Trim(sel_dup_dmpro_rs("reqno"))
'    sel_dmpro_qr(1) = Trim(sel_dup_dmpro_rs("PRO_CODE"))
'    Set sel_dmpro_rs = sel_dmpro_qr.OpenResultset(1, 3)
'    sel_dmpro_rs.MoveLast
'    'sel_dmpro_rs.MovePrevious
'    If Not sel_dmpro_rs.EOF Then
'        If IsNull(sel_dmpro_rs("REQNO")) = False Or Trim(sel_dmpro_rs("REQNO")) <> "" Then
'            reqno = Trim(sel_dmpro_rs("REQNO"))
'        Else
'            reqno = ""
'        End If
'        If IsNull(sel_dmpro_rs("PRO_CODE")) = False Or Trim(sel_dmpro_rs("PRO_CODE")) <> "" Then
'            PRO_CODE = Trim(sel_dmpro_rs("PRO_CODE"))
'        Else
'            PRO_CODE = ""
'        End If
'        If IsNull(sel_dmpro_rs("OBSERVALUE")) = False Or Trim(sel_dmpro_rs("OBSERVALUE")) <> "" Then
'            OBSERVALUE = Trim(sel_dmpro_rs("OBSERVALUE"))
'        Else
'            OBSERVALUE = ""
'        End If
'
'        If IsNull(sel_dmpro_rs("OBSER_TYPE")) = False Or Trim(sel_dmpro_rs("OBSER_TYPE")) <> "" Then
'            OBSER_TYPE = Trim(sel_dmpro_rs("OBSER_TYPE"))
'        Else
'            OBSER_TYPE = ""
'        End If
'
'        If IsNull(sel_dmpro_rs("SOURCE")) = False Or Trim(sel_dmpro_rs("SOURCE")) <> "" Then
'            Source = Trim(sel_dmpro_rs("SOURCE"))
'        Else
'            Source = ""
'        End If
'
'        If IsNull(sel_dmpro_rs("ENTRYBY")) = False Or Trim(sel_dmpro_rs("ENTRYBY")) <> "" Then
'            ENTRYBY = Trim(sel_dmpro_rs("ENTRYBY"))
'        Else
'            ENTRYBY = ""
'        End If
'
'        If IsNull(sel_dmpro_rs("REMARKS")) = False Or Trim(sel_dmpro_rs("REMARKS")) <> "" Then
'            REMARKS = Trim(sel_dmpro_rs("REMARKS"))
'        Else
'            REMARKS = ""
'        End If
'
'        If IsNull(sel_dmpro_rs("UOM")) = False Or Trim(sel_dmpro_rs("UOM")) <> "" Then
'            UOM = Trim(sel_dmpro_rs("UOM"))
'        Else
'            UOM = ""
'        End If
'
'        If IsNull(sel_dmpro_rs("NRANGE")) = False Or Trim(sel_dmpro_rs("NRANGE")) <> "" Then
'            NRANGE = Trim(sel_dmpro_rs("NRANGE"))
'        Else
'            NRANGE = ""
'        End If
'
'        If IsNull(sel_dmpro_rs("PVALUE")) = False Or Trim(sel_dmpro_rs("PVALUE")) <> "" Then
'            PVALUE = Trim(sel_dmpro_rs("PVALUE"))
'        Else
'            PVALUE = ""
'        End If
'    End If
'
'    del_dmpro_qr(0) = Trim(sel_dmpro_rs("REQNO"))
'    del_dmpro_qr(1) = Trim(sel_dmpro_rs("PRO_CODE"))
'    del_dmpro_qr.Execute
'
'    'reqno, PRO_CODE, OBSERVALUE, OBSER_TYPE, Source, ENTRYBY, REMARKS, UOM, NRANGE, PVALUE
'    INS_DMRRO_qr(0) = reqno
'    INS_DMRRO_qr(1) = PRO_CODE
'    INS_DMRRO_qr(2) = OBSERVALUE
'    INS_DMRRO_qr(3) = OBSER_TYPE
'    INS_DMRRO_qr(4) = Source
'    INS_DMRRO_qr(5) = ENTRYBY
'    INS_DMRRO_qr(6) = REMARKS
'    INS_DMRRO_qr(7) = UOM
'    INS_DMRRO_qr(8) = NRANGE
'    INS_DMRRO_qr(9) = PVALUE
'
'    INS_DMRRO_qr.Execute
'    sel_dup_dmpro_rs.MoveNext
'Next
''Wend
'
'MsgBox "Done!!!"
'End Sub
'
Private Sub Form_Load()

End Sub
