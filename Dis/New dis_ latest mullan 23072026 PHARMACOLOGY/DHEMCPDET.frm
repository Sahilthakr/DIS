VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form DHEMCPDET 
   Caption         =   "Form2"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8610
   LinkTopic       =   "Form2"
   ScaleHeight     =   8310
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton cmd_delete 
      Caption         =   "Delete Duplicates"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   8280
      Width           =   2655
   End
   Begin VB.CommandButton cmd_Show 
      Caption         =   "Show Duplicate records"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   8280
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7575
      Left            =   0
      TabIndex        =   3
      Top             =   480
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
      Left            =   2880
      TabIndex        =   5
      Top             =   9480
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "DHEMCPDET"
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
      Left            =   3840
      TabIndex        =   4
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "DHEMCPDET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim str As String
Dim sel_dup_dmpro_qr As rdoQuery
Dim sel_dup_dmpro_rs As rdoResultset

Dim sel_dmpro_qr As rdoQuery
Dim sel_dmpro_rs As rdoResultset

Dim del_dmpro_qr As rdoQuery

Dim INS_DMRRO_qr As rdoQuery

'Dim reqno, PRO_CODE, OBSERVALUE, OBSER_TYPE, Source, ENTRYBY, REMARKS, UOM, NRANGE, PVALUE As String

Dim REQNO, PROCCODE, OBSERVALUE, OBSERTYPE, CONTROL, ENTERBY, UOM, NRANGE, MACHINE, FLD1, FLD2
'REQNO PROCCODE  OBSERVALUE  OBSERTYPE  CONTROL ENTERBY  UOM NRANGE MACHINE FLD1 FLD2

Dim I As Long

Private Sub cmd_exit_Click()
End
End Sub

Private Sub Form_Load()
'str = "SELECT REQNO, PRO_CODE, COUNT(*) FROM DISLIB.DHEMCBCDET GROUP BY REQNO,PRO_CODE Having count(reqno) > 1"
str = "SELECT REQNO, PROCCODE, COUNT(*) FROM DISLIB.DHEMCPDET GROUP BY REQNO,PROCCODE Having count(reqno) > 1"
Set sel_dup_dmpro_qr = rdb.CreateQuery("", str)

'REQNO PRO_CODE  OBSERVALUE  OBSER_TYPE  SOURCE  ENTRYBY REMARKS UOM NRANGE PVALUE
'REQNO PROCCODE  OBSERVALUE  OBSERTYPE  CONTROL ENTERBY  UOM NRANGE MACHINE FLD1 FLD2

'str = "SELECT * FROM DISLIB.DHEMCBCDET where reqno=? AND PRO_CODE=?"
str = "SELECT * FROM DISLIB.DHEMCPDET where reqno=? AND PROCCODE=?"
Set sel_dmpro_qr = rdb.CreateQuery("", str)

'str = "delete from DISLIB.DHEMCBCDET where reqno=? AND PRO_CODE=?"
str = "delete from DISLIB.DHEMCPDET where reqno=? AND PROCCODE=?"
Set del_dmpro_qr = rdb.CreateQuery("", str)

'str = "INSERT INTO DISLIB.DHEMCBCDET VALUES(?,?,?,?,?,?,?,?,?,?)"
str = "INSERT INTO DISLIB.DHEMCPDET VALUES(?,?,?,?,?,?,?,?,?,?,?)"
Set INS_DMRRO_qr = rdb.CreateQuery("", str)
End Sub

Private Sub cmd_show_Click()
Set sel_dup_dmpro_rs = sel_dup_dmpro_qr.OpenResultset(1, 3)
MSFlexGrid1.Rows = sel_dup_dmpro_rs.RowCount + 1
MSFlexGrid1.Cols = 4

MSFlexGrid1.TextMatrix(0, 0) = "REQNO"
MSFlexGrid1.TextMatrix(0, 1) = "PROCCODE"
MSFlexGrid1.TextMatrix(0, 2) = "Count"

MSFlexGrid1.ColWidth(0) = 500
MSFlexGrid1.ColWidth(1) = 2000
MSFlexGrid1.ColWidth(2) = 1000
MSFlexGrid1.ColWidth(3) = 1000

For I = 1 To sel_dup_dmpro_rs.RowCount
    MSFlexGrid1.TextMatrix(I, 0) = I
    MSFlexGrid1.TextMatrix(I, 1) = sel_dup_dmpro_rs(0)
    MSFlexGrid1.TextMatrix(I, 2) = sel_dup_dmpro_rs(1)
    MSFlexGrid1.TextMatrix(I, 3) = sel_dup_dmpro_rs(2)
   sel_dup_dmpro_rs.MoveNext
Next I

End Sub

Private Sub cmd_delete_Click()
sel_dup_dmpro_rs.MoveFirst
For I = 1 To sel_dup_dmpro_rs.RowCount
'While Not sel_dup_dmpro_rs.EOF
    sel_dmpro_qr(0) = Trim(sel_dup_dmpro_rs("reqno"))
    sel_dmpro_qr(1) = Trim(sel_dup_dmpro_rs("PROCCODE"))
    Set sel_dmpro_rs = sel_dmpro_qr.OpenResultset(1, 3)
    sel_dmpro_rs.MoveLast
    'sel_dmpro_rs.MovePrevious
    If Not sel_dmpro_rs.EOF Then
        If IsNull(sel_dmpro_rs("REQNO")) = False Or Trim(sel_dmpro_rs("REQNO")) <> "" Then
            REQNO = Trim(sel_dmpro_rs("REQNO"))
        Else
            REQNO = ""
        End If
        If IsNull(sel_dmpro_rs("PROCCODE")) = False Or Trim(sel_dmpro_rs("PROCCODE")) <> "" Then
            PROCCODE = Trim(sel_dmpro_rs("PROCCODE"))
        Else
            PROCCODE = ""
        End If
        
        If IsNull(sel_dmpro_rs("OBSERVALUE")) = False Or Trim(sel_dmpro_rs("OBSERVALUE")) <> "" Then
            OBSERVALUE = Trim(sel_dmpro_rs("OBSERVALUE"))
        Else
            OBSERVALUE = ""
        End If

        If IsNull(sel_dmpro_rs("OBSERTYPE")) = False Or Trim(sel_dmpro_rs("OBSERTYPE")) <> "" Then
            OBSERTYPE = Trim(sel_dmpro_rs("OBSERTYPE"))
        Else
            OBSERTYPE = ""
        End If

        If IsNull(sel_dmpro_rs("CONTROL")) = False Or Trim(sel_dmpro_rs("CONTROL")) <> "" Then
            CONTROL = Trim(sel_dmpro_rs("CONTROL"))
        Else
            CONTROL = ""
        End If

        If IsNull(sel_dmpro_rs("ENTERBY")) = False Or Trim(sel_dmpro_rs("ENTERBY")) <> "" Then
            ENTERBY = Trim(sel_dmpro_rs("ENTERBY"))
        Else
            ENTERBY = ""
        End If

        If IsNull(sel_dmpro_rs("UOM")) = False Or Trim(sel_dmpro_rs("UOM")) <> "" Then
            UOM = Trim(sel_dmpro_rs("UOM"))
        Else
            UOM = ""
        End If

        If IsNull(sel_dmpro_rs("NRANGE")) = False Or Trim(sel_dmpro_rs("NRANGE")) <> "" Then
            NRANGE = Trim(sel_dmpro_rs("NRANGE"))
        Else
            NRANGE = ""
        End If

        If IsNull(sel_dmpro_rs("MACHINE")) = False Or Trim(sel_dmpro_rs("MACHINE")) <> "" Then
            MACHINE = Trim(sel_dmpro_rs("MACHINE"))
        Else
            MACHINE = ""
        End If
        

        If IsNull(sel_dmpro_rs("FLD1")) = False Or Trim(sel_dmpro_rs("FLD1")) <> "" Then
            FLD1 = Trim(sel_dmpro_rs("FLD1"))
        Else
            FLD1 = ""
        End If
        
        If IsNull(sel_dmpro_rs("FLD2")) = False Or Trim(sel_dmpro_rs("FLD2")) <> "" Then
            FLD2 = Trim(sel_dmpro_rs("FLD2"))
        Else
            FLD2 = ""
        End If
    End If

    del_dmpro_qr(0) = Trim(sel_dmpro_rs("REQNO"))
    del_dmpro_qr(1) = Trim(sel_dmpro_rs("PROCCODE"))
    del_dmpro_qr.Execute

    'reqno, PRO_CODE, OBSERVALUE, OBSER_TYPE, Source, ENTRYBY, REMARKS, UOM, NRANGE, PVALUE
    'REQNO PROCCODE  OBSERVALUE  OBSERTYPE  CONTROL ENTERBY  UOM NRANGE MACHINE FLD1 FLD2
    INS_DMRRO_qr(0) = REQNO
    INS_DMRRO_qr(1) = PROCCODE
    INS_DMRRO_qr(2) = OBSERVALUE
    INS_DMRRO_qr(3) = OBSERTYPE
    INS_DMRRO_qr(4) = CONTROL
    INS_DMRRO_qr(5) = ENTERBY
    INS_DMRRO_qr(6) = UOM
    INS_DMRRO_qr(7) = NRANGE
    INS_DMRRO_qr(8) = MACHINE
    INS_DMRRO_qr(9) = FLD1
    INS_DMRRO_qr(10) = FLD2
    INS_DMRRO_qr.Execute
    sel_dup_dmpro_rs.MoveNext
Next
'Wend

MsgBox "Done!!!"
End Sub


