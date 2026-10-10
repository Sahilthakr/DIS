VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Antibiotic 
   Caption         =   "Antibiotic Master"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10665
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_clear 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   7920
      TabIndex        =   18
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "&Save"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8640
      TabIndex        =   14
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8640
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmd_view 
      Caption         =   "&View"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox txt_anti 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   3720
      Width           =   5175
   End
   Begin VB.ComboBox cmb_modality 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1920
      Width           =   5175
   End
   Begin VB.ComboBox cmb_submodality 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2520
      TabIndex        =   7
      Top             =   2880
      Width           =   5175
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "&Back"
      Height          =   495
      Left            =   9360
      TabIndex        =   5
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmd_modify 
      Caption         =   "&Modify"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmd_delete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmd_add 
      Caption         =   "&New"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   6600
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid grid_anti 
      Height          =   1695
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "To Modify Click on Antibiotic Name"
      Top             =   4560
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   2990
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ForeColor       =   8421504
      ForeColorFixed  =   192
      FormatString    =   $"antibiotic_form.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8160
      TabIndex        =   17
      Top             =   840
      Width           =   420
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8160
      TabIndex        =   16
      Top             =   1200
      Width           =   420
   End
   Begin VB.Label lbl_microbiology 
      AutoSize        =   -1  'True
      Caption         =   "Antimicrobial Matser"
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
      Left            =   3120
      TabIndex        =   15
      Top             =   720
      Width           =   2625
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Antiobiotic Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   12
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Select Sub Modality"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   720
      TabIndex        =   11
      Top             =   2880
      Width           =   1710
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "Select Modality"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1080
      TabIndex        =   10
      Top             =   1920
      Width           =   1320
   End
End
Attribute VB_Name = "frm_Antibiotic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim str                 As String
Dim sel_moadlityqr      As rdo.rdoQuery
Dim sel_moadlityrs      As rdo.rdoResultset

Dim sel_submoadlityqr   As rdo.rdoQuery
Dim sel_submoadlityrs   As rdo.rdoResultset

Dim sel_allantiqr       As rdo.rdoQuery
Dim sel_allantirs       As rdo.rdoResultset

Dim sel_antiqr          As rdo.rdoQuery

Dim ins_anti_qr         As rdo.rdoQuery
Dim upd_anti_qr         As rdo.rdoQuery
Dim del_anti_qr         As rdo.rdoQuery
Dim new_flag            As Boolean
Dim mod_flag            As Boolean
Dim del_flag            As Boolean
Dim VIEW_FLAG           As Boolean
Dim srno                As Byte
Dim code                As Integer
'here

Private Sub cmb_modality_Click()
On Error GoTo MSG
sel_submoadlityqr(0) = cmb_modality.Text
Set sel_submoadlityrs = sel_submoadlityqr.OpenResultset(1, 3)
    If sel_submoadlityrs.RowCount > 0 Then
        cmb_submodality.clear
        While Not sel_submoadlityrs.EOF
            cmb_submodality.AddItem Trim(sel_submoadlityrs("submodali"))
            sel_submoadlityrs.MoveNext
        Wend
        sel_submoadlityrs.Close
    Else
        sel_submoadlityrs.Close
    End If
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmb_submodality_Click()
On Error GoTo MSG
   display
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_add_Click()
On Error GoTo MSG
new_flag = True
btnenbl "NEW"
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_back_Click()
On Error GoTo MSG
    'btnenbl "NEW"
    DIS_MAIN.show
    Me.Hide
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_clear_Click()
On Error GoTo MSG
btnenbl "DELETE"
txt_anti.Text = ""
cmb_modality.ListIndex = -1
cmb_submodality.Text = ""
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_delete_Click()
On Error GoTo MSG
    del_anti_qr(0) = code
    del_anti_qr.Execute
    display
    MsgBox "Record Deleted Successfully", vbOKOnly
    btnenbl "DELETE"
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_modify_Click()
On Error GoTo MSG
cmd_modify.Enabled = False
cmd_delete.Enabled = False
cmd_save.Enabled = True
mod_flag = True
Exit Sub
MSG:
    MsgBox Err.Description

End Sub

Private Sub cmd_save_Click()
On Error GoTo MSG
    If new_flag Then
        If Len(txt_anti.Text) > 2 And Len(cmb_modality.Text) > 2 And Len(cmb_submodality.Text) > 2 Then
            srno = grid_anti.TextMatrix(grid_anti.Rows - 1, 0)
            srno = srno + 1
            ins_anti_qr(0) = srno
            ins_anti_qr(1) = txt_anti.Text
            ins_anti_qr(2) = cmb_modality.Text
            ins_anti_qr(3) = cmb_submodality.Text
            ins_anti_qr.Execute
            display
            new_flag = False
            MsgBox "Record Inserted Successfully", vbOKOnly
            btnenbl "SAVE"
            Exit Sub
         Else
            MsgBox "Enter Proper Details", vbOKOnly
            btnenbl "SAVE"
         End If
    ElseIf mod_flag Then
        upd_anti_qr(0) = txt_anti.Text
        upd_anti_qr(1) = cmb_modality
        upd_anti_qr(2) = cmb_submodality
        upd_anti_qr(3) = code
        upd_anti_qr.Execute
        display
        mod_flag = False
        code = 0
        MsgBox "Record Updated Successfully", vbOKOnly
        btnenbl "SAVE"
        Exit Sub
    End If
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub cmd_view_Click()
On Error GoTo MSG
VIEW_FLAG = True
btnenbl "VIEW"
Exit Sub
MSG:
    MsgBox Err.Description

End Sub

Private Sub Form_Load()
Dim str As String
On Error GoTo MSG
    datetime
    txt_date.Text = Format(DATE11, "DD/MM/YYYY")
    txt_time.Text = time1
    mod_flag = False
    VIEW_FLAG = False
    new_flag = False
    btnenbl "DELETE"
    grid_anti.Cols = 4
    '***********************************************************************
                        'query definition
    '***********************************************************************
    
    'to select modality
    str = "select * from rislib.rmodality WHERE DEPT_CODE=?"
    Set sel_moadlityqr = rdb.CreateQuery("Anti_select1", str)
    
    'to select sub modality
    str = "select distinct(submodali) from dislib.dmicbantib WHERE modality=?"
    Set sel_submoadlityqr = rdb.CreateQuery("Anti_select2", str)
    
    'to select fresh record from table
    str = "select antibcode,antibdesc,modality,submodali from dislib.dmicbantib "
    Set sel_allantiqr = rdb.CreateQuery("Anti_select3", str)
    
    'to select fresh record for particular modality
    str = "select antibcode,antibdesc,modality,submodali from dislib.dmicbantib  where modality=? and submodali=?"
    Set sel_antiqr = rdb.CreateQuery("Anti_select4", str)
    
    'to insert data in table
    str = "insert into dislib.dmicbantib(antibcode,antibdesc,modality,submodali)values(?,?,?,?)"
    Set ins_anti_qr = rdb.CreateQuery("Anti_insert1", str)
    
    'to update data in table
    str = "update dislib.dmicbantib set antibdesc=?,modality=?,submodali=? where antibcode=?"
    Set upd_anti_qr = rdb.CreateQuery("Anti_update1", str)
    
    'to dalete data in table
    str = "delete from dislib.dmicbantib where antibcode=?"
    Set del_anti_qr = rdb.CreateQuery("Anti_delete1", str)
    
    'to load modality to combo
    sel_moadlityqr(0) = "GZZ"
    Set sel_moadlityrs = sel_moadlityqr.OpenResultset(1, 3)
       If sel_moadlityrs.RowCount > 0 Then
            cmb_modality.clear
            While Not sel_moadlityrs.EOF
                cmb_modality.AddItem Trim(sel_moadlityrs("modlitydes"))
                sel_moadlityrs.MoveNext
            Wend
            sel_moadlityrs.Close
        Else
            sel_moadlityrs.Close
        End If
    display
    Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub grid_anti_Click()
On Error GoTo MSG
    If grid_anti.col = 1 Or grid_anti.col = 2 Then
        txt_anti.Text = Trim(grid_anti.TextMatrix(grid_anti.row, 1))
        code = CInt(CStr(grid_anti.TextMatrix(grid_anti.row, 0)))
    End If
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
Private Sub txt_anti_LOSTFOCUS()
On Error GoTo MSG
If Len(txt_anti.Text) > 2 Then
    txt_anti.Text = UCase(txt_anti.Text)
Else
    MsgBox "Please Enter Proper Antimicrobial Name", vbOKOnly
End If
Exit Sub
MSG:
    MsgBox Err.Description

End Sub



Private Sub display(Optional show As Byte)
On Error GoTo MSG
grid_anti.Cols = 4
If show = 5 Then
    sel_antiqr(0) = cmb_modality.Text
    sel_antiqr(1) = cmb_submodality.Text
    Set sel_allantirs = sel_antiqr.OpenResultset(1, 3)
Else
    Set sel_allantirs = sel_allantiqr.OpenResultset(1, 3)
End If
        srno = 0
        If Not sel_allantirs.EOF Then
            grid_anti.Rows = sel_allantirs.RowCount + 1
            grid_anti.clear
            grid_anti.FormatString = "Antibiotic Code |Antibiotic Name                                       |Modality                       |Sub Modality                   "
            While Not sel_allantirs.EOF
                srno = srno + 1
                grid_anti.TextMatrix(srno, 0) = sel_allantirs("antibcode")
                grid_anti.TextMatrix(srno, 1) = sel_allantirs("antibdesc")
                grid_anti.TextMatrix(srno, 2) = sel_allantirs("modality")
                grid_anti.TextMatrix(srno, 3) = sel_allantirs("submodali")
                sel_allantirs.MoveNext
            Wend
            sel_allantirs.Close
        Else
            sel_allantirs.Close
        End If
Exit Sub
MSG:
    MsgBox Err.Description
End Sub

Private Sub btnenbl(cap As String)
On Error GoTo MSG
If cap = "NEW" Then
    cmd_save.Enabled = True
    cmd_back.Enabled = True
    cmd_add.Enabled = False
    cmd_modify.Enabled = False
    cmd_delete.Enabled = False
    cmd_view.Enabled = False
ElseIf cap = "VIEW" Then
    cmd_save.Enabled = False
    cmd_back.Enabled = True
    cmd_add.Enabled = False
    cmd_modify.Enabled = True
    cmd_view.Enabled = False
    cmd_delete.Enabled = True
ElseIf cap = "SAVE" Then
    cmd_save.Enabled = False
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    cmd_add.Enabled = True
    cmd_view.Enabled = True
    cmd_delete.Enabled = False
ElseIf cap = "DELETE" Then
    cmd_save.Enabled = False
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    cmd_add.Enabled = True
    cmd_view.Enabled = True
    cmd_delete.Enabled = False
End If
Exit Sub
MSG:
    MsgBox Err.Description
End Sub
