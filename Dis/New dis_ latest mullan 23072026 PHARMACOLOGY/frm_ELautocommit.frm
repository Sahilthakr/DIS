VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_ELautocommit 
   Caption         =   "Form2"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12540
   LinkTopic       =   "Form2"
   ScaleHeight     =   6060
   ScaleWidth      =   12540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton commit_btn 
      Caption         =   "&COMMIT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   6375
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Text            =   "BIOCHEMISTRY$FEC"
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txt_modalitycd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox combo_modality 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_ELautocommit.frx":0000
         Left            =   1920
         List            =   "frm_ELautocommit.frx":0002
         TabIndex        =   2
         Top             =   960
         Width           =   2895
      End
      Begin MSMask.MaskEdBox MaskEd_todt 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Modality"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "AUTOCOMMIT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label frm_ELautocommit 
      Caption         =   "03/11/2021"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frm_ELautocommit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l, i As Integer
Dim str As String
Dim SEL_MOD_QR As rdoQuery
Dim SEL_MOD_RS As rdoResultset
Dim SEL_MOD_QR1 As rdoQuery
Dim SEL_MOD_RS1 As rdoResultset
Dim UPDREQ As rdoQuery

'*********** Shraddha 30/06/09 **************
Dim upd_rtimelogqr As rdoQuery
Dim sel_rreq_qr As rdoQuery
Dim sel_rreq_rs As rdoResultset
Dim REQNO As String
'********************************************

Dim to_date As Date
Dim cur_date, pre_date As Date
Dim c_dt As String

Private Sub cmd_back_Click()
Me.Hide
DIS_MAIN.show
End Sub

Private Sub combo_modality_LostFocus()
If combo_modality.Text = "" Then
    MsgBox "Please select Modality", vbCritical
    Exit Sub
End If

str = "SELECT * FROM RISLIB.RMODALITY WHERE MODLITYDES = ?"
Set SEL_MOD_QR1 = rdb.CreateQuery("", str)
SEL_MOD_QR1(0) = combo_modality.Text
Set SEL_MOD_RS1 = SEL_MOD_QR1.OpenResultset(1, 3)
txt_modalitycd.Text = Trim(SEL_MOD_RS1("MODALITYCD"))
End Sub


Private Sub commit_btn_Click()

'*********** Shraddha 29/06/09 **********
Dim tdt As Date
Dim TTIME
Dim r As Integer
'*******************************

'Check whether to date is empty or not
If MaskEd_todt.Text = "__/__/____" Then
    MsgBox "Please Enter To Date", vbCritical
    Exit Sub
End If

If Not IsDate(MaskEd_todt.Text) Then
    MsgBox "Invalid Date", vbCritical
    MaskEd_todt.SetFocus
    Exit Sub
End If

If combo_modality.Text = "" Then
    MsgBox "Please select Modality", vbCritical
    Exit Sub
End If

to_date = Format(MaskEd_todt.Text, "yyyy-mm-dd")
'cur_date = DateAdd("d", -4, Date)
'c_dt = cur_date

'If to_date > cur_date Then
'    MsgBox "To date can not be greater than " + c_dt + vbCrLf + "Please re-enter date"
'    Exit Sub
'End If

'str = "update rislib.rrequisitn set reqnstat='F',commitflag='Y' "
'str = str + "where  reqnstat='T' and  FREPORTDT <= ?"
'str = str + "and dept_code='FZZ' and modalitycd=?"
'
'Set UPDREQ = rdb.CreateQuery("", str)
'
'UPDREQ(0) = to_date
'UPDREQ(1) = txt_modalitycd
'
'UPDREQ.Execute
'Dim n As Integer
'n = UPDREQ.RowsAffected
'**************** Shraddha Pandit 30/06/2009 ***********************
datetime
'MsgBox DATE11
'MsgBox time1
tdt = Format(DATE11, "yyyy-mm-dd")
TTIME = time1

str = "select * from rislib.rrequisitn where reqnstat='T' and  FREPORTDT <= ? and dept_code='FEC' and modalitycd=?"
Set sel_rreq_qr = rdb.CreateQuery("", str)

sel_rreq_qr(0) = to_date
sel_rreq_qr(1) = txt_modalitycd
Set sel_rreq_rs = sel_rreq_qr.OpenResultset(1, 3)

str = "update rislib.rtimelog set COMREPDATE='" & Format(DATE11, "yyyy-mm-dd") & "', COMREPTIME='" & time1 & "'where reqnno=?"
Set upd_rtimelogqr = rdb.CreateQuery("", str)

While Not sel_rreq_rs.EOF
    REQNO = sel_rreq_rs("reqnno")
    upd_rtimelogqr(0) = REQNO
    upd_rtimelogqr.Execute
    sel_rreq_rs.MoveNext
Wend

'********************************************************************

str = "update rislib.rrequisitn set reqnstat='F',commitflag='Y' "
str = str + "where  reqnstat='T' and  FREPORTDT <= ?"
str = str + "and dept_code='FEC' and modalitycd=?"

Set UPDREQ = rdb.CreateQuery("", str)

UPDREQ(0) = to_date
UPDREQ(1) = txt_modalitycd

UPDREQ.Execute
Dim n As Integer
n = UPDREQ.RowsAffected

MsgBox "Records successfully committed"

End Sub

Private Sub Form_Load()

cur_date = Format(Date, "yyyy-mm-dd")
str = "SELECT * FROM RISLIB.RMODALITY WHERE DEPT_CODE = ?"
Set SEL_MOD_QR = rdb.CreateQuery("", str)
SEL_MOD_QR(0) = "FEC"
Set SEL_MOD_RS = SEL_MOD_QR.OpenResultset(1, 3)
l = SEL_MOD_RS.RowCount
SEL_MOD_RS.MoveFirst
While Not SEL_MOD_RS.EOF
    'combo_modality.AddItem (SEL_MOD_RS("MODLITYDES"))
    SEL_MOD_RS.MoveNext
Wend

combo_modality.AddItem ("EMERGENCY LAB")

'MaskEd_todt.Text = Date
datetime
MaskEd_todt = DATE11
'MaskEd_todt = Format(DATE11, "dd-mm-yyyy")
End Sub

