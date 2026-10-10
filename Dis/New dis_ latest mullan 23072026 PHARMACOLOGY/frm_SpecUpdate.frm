VERSION 5.00
Begin VB.Form frm_SpecUpdate 
   BackColor       =   &H00808080&
   Caption         =   "Specimen Update"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "Go to Main Menu"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmd_update 
      Caption         =   "Update"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox cmb_specimen 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txt_reqno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      MaxLength       =   13
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "13 Digit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Specimen"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Requisition No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frm_SpecUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sel_SpecimenListQR As rdoQuery
Dim sel_SpecimenListRS As rdoResultset


Dim sel_SpecimenDQR As rdoQuery
Dim sel_SpecimenDRS As rdoResultset


Dim UPDRREQ_QR As rdoQuery
Dim str As String
Dim REQ1 As String


Private Sub cmd_back_Click()
Me.Hide
DIS_MAIN.show

End Sub

Private Sub cmd_update_Click()
If Len(Trim(txt_reqno.Text)) = 13 Then

UPDRREQ_QR(0) = cmb_specimen.Text
UPDRREQ_QR(1) = txt_reqno.Text
UPDRREQ_QR.Execute
If UPDRREQ_QR.RowsAffected = 1 Then
MsgBox "Record Updated"
Else
MsgBox "Please check Once Again"
End If
Else
MsgBox "Please Enter Valid Requisition No.(13 Digit)"
End If
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
logonfrm.Main
datetime
str = "UPDATE DISLIB.DMSPONANTI SET SPECIMEN=? WHERE REQNO=?"
Set UPDRREQ_QR = rdb.CreateQuery("", str)



str = "SELECT SAMPLENAME FROM DISLIB.MICROSAMPL"
Set sel_SpecimenListQR = rdb.CreateQuery("", str)
Set sel_SpecimenListRS = sel_SpecimenListQR.OpenResultset(1, 3)
If sel_SpecimenListRS.EOF = False Then
While Not sel_SpecimenListRS.EOF
cmb_specimen.AddItem (Trim(sel_SpecimenListRS(0)))
sel_SpecimenListRS.MoveNext
Wend
End If
'txt_reqno.SetFocus
End Sub



Private Sub txt_reqno_LostFocus()
If Len(Trim(txt_reqno.Text)) = 13 Then
    str = "SELECT SPECIMEN FROM DISLIB.DMSPONANTI WHERE REQNO=?"
    Set sel_SpecimenDQR = rdb.CreateQuery("", str)
    sel_SpecimenDQR(0) = Trim(txt_reqno.Text)
    Set sel_SpecimenDRS = sel_SpecimenDQR.OpenResultset(1, 3)
    If sel_SpecimenDRS.EOF = False Then
   ' While Not sel_SpecimenListRS.EOF
    cmb_specimen.Text = (Trim(sel_SpecimenDRS(0)))
    'sel_SpecimenListRS.MoveNext
   ' Wend
    End If
End If

End Sub
