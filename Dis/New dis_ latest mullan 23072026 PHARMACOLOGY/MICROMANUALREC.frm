VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MICROMANUALREC 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Microbiology Manual Receiving"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   7920
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   615
      Left            =   6720
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton CMD_UPDMICROPATH 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   7560
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox TXT_PATHNO 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   5160
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Width           =   1140
   End
   Begin MSFlexGridLib.MSFlexGrid MSK_MANUALREC 
      Height          =   3855
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   7
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "MICROBIOLOGY MANUAL RECEIVING UPDATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   7920
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbl_datetime 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7320
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "MICROMANUALREC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sel_MICROPATHqr As rdoQuery
Dim sel_MICROPATHrs As rdoResultset
Dim SEL_STR As String

Private Sub Command1_Click()
Unload Me
DIS_MAIN.show
End Sub

Private Sub CMD_UPDMICROPATH_Click()
a = MsgBox("Do you want to update", vbYesNo, "PATHNO")
If a = 6 Then
    Dim UPDMICROPATH_QR As rdoQuery, str As String
    Dim RW As Integer
    
    str = "UPDATE DISLIB.MICROPATH SET PATHNO=? WHERE MODALITY=? AND TYPE=? AND STATUS=?"
    Set UPDMICROPATH_QR = rdb.CreateQuery("", str)
    
    For RW = 2 To MSK_MANUALREC.Rows - 1
    UPDMICROPATH_QR(0) = MSK_MANUALREC.TextMatrix(RW, 5)
    UPDMICROPATH_QR(1) = MSK_MANUALREC.TextMatrix(RW, 1)
    UPDMICROPATH_QR(2) = MSK_MANUALREC.TextMatrix(RW, 6)
    UPDMICROPATH_QR(3) = MSK_MANUALREC.TextMatrix(RW, 2)
    UPDMICROPATH_QR.Execute
    Next
    
    MsgBox "Record Updated !!!"
    
    
End If
End Sub

Private Sub Form_Load()
datetime
lbl_datetime.Caption = DATE11 ' & time1
Dim ROWVAL As String
MSK_MANUALREC.row = 0
ROWVAL = "SrNo.|Modality | Status    | Prefix   | Pathno        |                       | Type"
MSK_MANUALREC.FormatString = ROWVAL

Dim K As Integer, i As Integer
SEL_STR = "SELECT * FROM DISLIB.MICROPATH"
Set sel_MICROPATHqr = rdb.CreateQuery("", SEL_STR)
Set sel_MICROPATHrs = sel_MICROPATHqr.OpenResultset(1, 3)
If sel_MICROPATHrs.RowCount > 0 Then
'MSK_MANUALREC.DataSource =
MSK_MANUALREC.row = 1

    For K = 0 To sel_MICROPATHrs.RowCount - 1
    i = K + 1
    MSK_MANUALREC.row = i
    
    ROWVAL = i & vbTab & Trim(sel_MICROPATHrs("MODALITY")) & vbTab & Trim(sel_MICROPATHrs("STATUS")) & vbTab & Trim(sel_MICROPATHrs("PREFIX1")) & vbTab & Trim(sel_MICROPATHrs("PATHNO")) & vbTab & "" & vbTab & Trim(sel_MICROPATHrs("TYPE"))
        MSK_MANUALREC.AddItem ROWVAL
        sel_MICROPATHrs.MoveNext
    Next
End If
'MSK_MANUALREC.row = 1
'MSK_MANUALREC.col = 5
'MSK_MANUALREC.CellWidth(5) = 500

End Sub

Private Sub MSK_MANUALREC_Click()
If MSK_MANUALREC.col = 5 Then

        txt_pathno.Visible = True
        txt_pathno.Top = MSK_MANUALREC.CellTop + MSK_MANUALREC.Top
        txt_pathno.Left = MSK_MANUALREC.CellLeft + MSK_MANUALREC.Left
        txt_pathno.Width = MSK_MANUALREC.CellWidth + 10
        txt_pathno.Text = MSK_MANUALREC.TextMatrix(MSK_MANUALREC.row, 4)
        txt_pathno.SelStart = 6
        txt_pathno.SelLength = 4
End If
End Sub

Private Sub TXT_PATHNO_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
If Trim(MSK_MANUALREC.TextMatrix(MSK_MANUALREC.row, 4)) <> Trim(txt_pathno.Text) Then
Dim k1 As Integer
For k1 = 0 To 5
MSK_MANUALREC.col = k1
MSK_MANUALREC.CellBackColor = AQUA
Next
End If


MSK_MANUALREC.TextMatrix(MSK_MANUALREC.row, 5) = txt_pathno.Text

If MSK_MANUALREC.row = MSK_MANUALREC.Rows - 1 Then
txt_pathno.Visible = False
Exit Sub
End If

MSK_MANUALREC.row = MSK_MANUALREC.row + 1
        txt_pathno.Visible = True
        txt_pathno.Top = MSK_MANUALREC.CellTop + MSK_MANUALREC.Top
        txt_pathno.Left = MSK_MANUALREC.CellLeft + MSK_MANUALREC.Left
        txt_pathno.Width = MSK_MANUALREC.CellWidth + 10
        txt_pathno.Text = MSK_MANUALREC.TextMatrix(MSK_MANUALREC.row, 4)
        txt_pathno.SelStart = 6
        txt_pathno.SelLength = 4
'        TXT_PATHNO.SetFocus
End If
End Sub
