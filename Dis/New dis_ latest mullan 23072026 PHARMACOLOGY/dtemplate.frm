VERSION 5.00
Begin VB.Form dtemplate 
   BackColor       =   &H00E0E0E0&
   Caption         =   "DTEMPLATE"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12630
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
   ScaleHeight     =   8580
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_procdesc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7680
      MaxLength       =   90
      TabIndex        =   15
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   840
      TabIndex        =   16
      Top             =   2400
      Width           =   10455
      Begin VB.TextBox txt_proccode 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1920
         TabIndex        =   23
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txt_lowval 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1920
         TabIndex        =   22
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txt_highval 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6840
         TabIndex        =   21
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txt_uom 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txt_sex 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6840
         TabIndex        =   19
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txt_lpvalue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txt_hpvalue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6840
         TabIndex        =   17
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lbl_proccode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procedure Code"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label lbl_lowval 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Low Value"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   840
         TabIndex        =   30
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lbl_highval 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "High value"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5760
         TabIndex        =   29
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label lbl_uom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit of Measurment"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   1665
      End
      Begin VB.Label lbl_sex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6240
         TabIndex        =   27
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label LBL_PLOWVALUE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Low Panic Value"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   2280
         Width           =   1440
      End
      Begin VB.Label lbl_hpvalue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "High Panic Value"
         Height          =   195
         Left            =   5160
         TabIndex        =   25
         Top             =   2280
         Width           =   1485
      End
      Begin VB.Label lbl_procdesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procedure Description"
         Height          =   195
         Left            =   4680
         TabIndex        =   24
         Top             =   600
         Width           =   1905
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   3135
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   10455
      End
   End
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9960
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txt_time 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9960
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame buttons 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   840
      TabIndex        =   7
      Top             =   5880
      Width           =   10455
      Begin VB.CommandButton cmd_next 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ne&xt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7830
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_prev 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Previous"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3975
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_back 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9120
         MouseIcon       =   "dtemplate.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_new 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MouseIcon       =   "dtemplate.frx":030A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_view 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1405
         MouseIcon       =   "dtemplate.frx":0614
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_modify 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2690
         MouseIcon       =   "dtemplate.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5260
         MouseIcon       =   "dtemplate.frx":0C28
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_cancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6545
         MouseIcon       =   "dtemplate.frx":0F32
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      Left            =   960
      TabIndex        =   12
      Top             =   240
      Width           =   8595
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9120
      TabIndex        =   11
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label lbl_time 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Time"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9120
      TabIndex        =   10
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PROCEDURE MASTER-REPORTING "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   960
      Width           =   4215
   End
End
Attribute VB_Name = "dtemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim insert_qr As rdoQuery
Dim VIEW_qr As rdoQuery
Dim PRDESC_qr As rdoQuery
Dim PRDESC_rs As rdoResultset
Dim VIEW_RS As rdoResultset
Dim VIEWF_qr As rdoQuery
Dim VIEWF_RS As rdoResultset
Dim viewflag As Boolean
Dim NEWFLAG As Boolean
Dim UPDATE_QR As rdoQuery
Dim UPDATE_rs As rdoResultset
Dim a As Integer

Private Sub cmd_back_Click()
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_modify_Click()
btnenabl ("MODIFY")
FLDENABL ("MODIFY")
End Sub

Private Sub CMD_NEXT_Click()
'Add by Milind date 12/01/06
Dim row, i As Integer
row = VIEWF_RS.RowCount

    For i = 1 To row
        If VIEWF_RS.BOF = True Then
            VIEWF_RS.MoveNext
        End If
        If VIEWF_RS(0) = txt_proccode.Text Then
            If VIEWF_RS.EOF = False Then
                VIEWF_RS.MoveNext
            End If
            Exit For
            Else
                VIEWF_RS.MoveNext
        End If
    Next
         
    
    If VIEWF_RS.RowCount > 0 And VIEWF_RS.EOF = False Then
    txt_proccode.Text = VIEWF_RS(0)
    'change in proc. description dt.13/01/06 milind
    PRDESC_qr(0) = Trim(txt_proccode.Text)
    Set PRDESC_rs = PRDESC_qr.OpenResultset(2, 3)
    If PRDESC_rs.RowCount <> 0 Then
        txt_procdesc.Text = PRDESC_rs(0)
    End If
    txt_sex.Text = VIEWF_RS(1)
    txt_lowval.Text = VIEWF_RS(2)
    txt_highval.Text = VIEWF_RS(3)
    txt_uom.Text = VIEWF_RS(4)
    txt_lpvalue.Text = VIEWF_RS(5)
    txt_hpvalue.Text = VIEWF_RS(6)
    Else
    MsgBox " No record to display"
    clear
    btnenabl ("PREVIOUS")
    End If
End Sub

Private Sub cmd_prev_Click()
'Add by Milind date 12/01/06
 
    Dim row, i As Integer
    row = VIEWF_RS.RowCount
    
    If VIEWF_RS.BOF = False Then
        For i = 1 To row
            If VIEWF_RS(0) = txt_proccode.Text Then
                If VIEWF_RS.BOF = False Then
                    VIEWF_RS.MovePrevious
                End If
                Exit For
            Else
                VIEWF_RS.MoveNext
            End If
        Next
    End If
    
    If VIEWF_RS.RowCount > 0 And Not VIEWF_RS.BOF Then
    txt_proccode.Text = VIEWF_RS(0)
    
    'change in proc. description dt.13/01/06 milind
    PRDESC_qr(0) = Trim(txt_proccode.Text)
    Set PRDESC_rs = PRDESC_qr.OpenResultset(2, 3)
    If PRDESC_rs.RowCount <> 0 Then
        txt_procdesc.Text = PRDESC_rs(0)
        Else
        txt_procdesc.Text = ""
        End If
    
    txt_sex.Text = VIEWF_RS(1)
    txt_lowval.Text = VIEWF_RS(2)
    txt_highval.Text = VIEWF_RS(3)
    txt_uom.Text = VIEWF_RS(4)
    txt_lpvalue.Text = VIEWF_RS(5)
    txt_hpvalue.Text = VIEWF_RS(6)
    Else
    MsgBox " This is the first record"
    btnenabl ("PREVIOUS")
    End If
    Exit Sub

End Sub

Private Sub cmd_save_Click()
If NEWFLAG Then
btnenabl ("SAVE")
insert_qr(0) = txt_proccode.Text
insert_qr(1) = txt_sex.Text
insert_qr(2) = txt_lowval.Text
insert_qr(3) = txt_highval.Text
insert_qr(4) = txt_uom.Text
insert_qr(5) = txt_lpvalue.Text
insert_qr(6) = txt_hpvalue.Text
insert_qr.Execute
MsgBox " Record saved Successfully"
clear
ElseIf viewflag Then

UPDATE_QR(0) = txt_sex.Text
UPDATE_QR(1) = txt_lowval.Text
UPDATE_QR(2) = txt_highval.Text
UPDATE_QR(3) = txt_uom.Text
UPDATE_QR(4) = txt_lpvalue.Text
UPDATE_QR(5) = txt_hpvalue.Text
UPDATE_QR(6) = txt_proccode.Text
UPDATE_QR(7) = txt_sex.Text

UPDATE_QR.Execute
MsgBox " Modified Record  Successfully"
clear
End If
End Sub

Private Sub Form_Load()
'Main
clear
datetime
    txt_date.Text = DATE11
    txt_time.Text = time1
btnenabl ("PREVIOUS")
NEWFLAG = False
viewflag = False

' Insert New Record
str = "insert into dislib.dtemplet (pro_code,sex,lowvalue,highvalue,uom,lowpanic,highpanic) values(?,?,?,?,?,?,?)"
Set insert_qr = rdb.CreateQuery("insert_qr1", str)
str = "SELECT * FROM DISLIB.DTEMPLET WHERE PRO_CODE =?"
Set VIEW_qr = rdb.CreateQuery("VIEW_QR1", str)
VIEW_qr(0) = "A"

str = "SELECT * FROM DISLIB.DTEMPLET"
Set VIEWF_qr = rdb.CreateQuery("VIEWF_QR1", str)
Set VIEWF_RS = VIEWF_qr.OpenResultset(2, 3)

'To display Procedure Description
str = "select SER_DET from patoolslib.ma101 where patoolslib.ma101.ser_code=?"
Set PRDESC_qr = rdb.CreateQuery("PROCDESC_qr1", str)
'Set PRDESC_rs = PRDESC_qr.OpenResultset(2, 3)


' Update Modified Record
str = "UPDATE dislib.dtemplet set  sex = ?, lowvalue = ?, highvalue = ? ,uom = ?,lowpanic=?,highpanic=? WHERE PRO_CODE=? and sex=?"
Set UPDATE_QR = rdb.CreateQuery("UPDATE_QR1", str)

End Sub

Public Sub btnenabl(cap As String)
If cap = "NEW" Then
    cmd_save.Enabled = True
    CMD_CANCEL.Enabled = True
    cmd_back.Enabled = True
    CMD_NEW.Enabled = False
    cmd_modify.Enabled = False
    cmd_view.Enabled = False
    
ElseIf cap = "VIEW" Then
    CMD_CANCEL.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = True
    CMD_NEW.Enabled = False
    cmd_save.Enabled = False
    cmd_view.Enabled = False
    cmd_next.Enabled = True
    cmd_prev.Enabled = True
    
ElseIf cap = "SAVE" Then
    CMD_CANCEL.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_view.Enabled = True
    
ElseIf cap = "PREVIOUS" Then
    CMD_CANCEL.Enabled = True
    cmd_back.Enabled = True
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = True
    cmd_save.Enabled = False
    cmd_view.Enabled = True
    cmd_prev.Enabled = False
    cmd_next.Enabled = False
  
ElseIf cap = "MODIFY" Then
    CMD_CANCEL.Enabled = True
    cmd_back.Enabled = False
    cmd_modify.Enabled = False
    CMD_NEW.Enabled = False
    cmd_save.Enabled = True
    cmd_view.Enabled = False
 End If
End Sub

Public Sub FLDENABL(str1 As String)
If str1 = "PREVIOUS" Then
      txt_proccode.Enabled = False
      txt_sex.Enabled = False
      txt_lowval.Enabled = False
      txt_highval.Enabled = False
      txt_uom.Enabled = False
End If
If str1 = "NEW" Then
      txt_proccode.Enabled = True
      txt_sex.Enabled = True
      txt_lowval.Enabled = True
      txt_highval.Enabled = True
      txt_uom.Enabled = True
End If
If str1 = "VIEW" Then
      txt_proccode.Enabled = True
      txt_sex.Enabled = False
      txt_lowval.Enabled = False
      txt_highval.Enabled = False
      txt_uom.Enabled = False
      txt_lpvalue.Enabled = False
      txt_hpvalue.Enabled = False
End If
If str1 = "CANCEL" Then
      txt_proccode.Enabled = False
      txt_procdesc.Text = ""
      txt_sex.Enabled = False
      txt_lowval.Enabled = False
      txt_highval.Enabled = False
      txt_uom.Enabled = False
      txt_hpvalue.Enabled = False
      txt_lpvalue.Enabled = False
End If
If str1 = "MODIFY" Then
      txt_proccode.Enabled = False
      txt_sex.Enabled = True
      txt_lowval.Enabled = True
      txt_highval.Enabled = True
      txt_uom.Enabled = True
      txt_hpvalue.Enabled = True
      txt_lpvalue.Enabled = True
End If
End Sub

Private Sub cmd_view_Click()
viewflag = True
btnenabl ("VIEW")
FLDENABL ("VIEW")
txt_proccode.SetFocus
End Sub
Private Sub cmd_new_Click()
btnenabl ("NEW")
FLDENABL ("NEW")
clear
txt_proccode.SetFocus
NEWFLAG = True
End Sub
Private Sub cmd_cancel_Click()
clear
btnenabl ("PREVIOUS")
FLDENABL ("CANCEL")

End Sub

Public Sub clear()
txt_highval.Text = " "
txt_lowval.Text = " "
txt_proccode.Text = " "
txt_sex.Text = " "
txt_uom.Text = " "
txt_hpvalue.Text = ""
txt_lpvalue.Text = ""
NEWFLAG = False
viewflag = False
End Sub



Private Sub txt_proccode_LostFocus()
Dim str As String
str = Trim(UCase(txt_proccode.Text))
txt_proccode.Text = str

If viewflag Then
    VIEW_qr(0) = Trim(txt_proccode.Text)
    Set VIEW_RS = VIEW_qr.OpenResultset(1, 3)
    PRDESC_qr(0) = Trim(txt_proccode.Text)
    Set PRDESC_rs = PRDESC_qr.OpenResultset(1, 3)
    If PRDESC_rs.RowCount <> 0 Then
        txt_procdesc.Text = PRDESC_rs(0)
    End If
    
    If VIEW_RS.RowCount > 0 Then
    
    txt_sex.Text = VIEW_RS(1)
    txt_lowval.Text = VIEW_RS(2)
    txt_highval.Text = VIEW_RS(3)
    txt_uom.Text = VIEW_RS(4)
    txt_lpvalue.Text = VIEW_RS(5)
    txt_hpvalue.Text = VIEW_RS(6)
    Else
    MsgBox " No record to display"
    clear
    btnenabl ("PREVIOUS")
    End If
End If
End Sub

