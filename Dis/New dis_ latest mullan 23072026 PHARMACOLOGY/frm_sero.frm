VERSION 5.00
Begin VB.Form frm_sero 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11355
   LinkTopic       =   "Form2"
   ScaleHeight     =   8010
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   3480
      TabIndex        =   0
      Top             =   2640
      Width           =   5655
      Begin VB.CommandButton cmd_back 
         Caption         =   "BACK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   4
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cmd_clickit 
         Caption         =   "REFRESH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   3
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox TXT_REQNO 
         Height          =   405
         Left            =   2640
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lbl_reqno 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Requisition No."
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
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   960
         Width           =   1440
      End
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
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   2220
      TabIndex        =   5
      Top             =   1440
      Width           =   8595
   End
End
Attribute VB_Name = "frm_sero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UPDRREQ_QR As rdoQuery
Dim del_micro As rdoQuery
Dim DEL_COMMENT As rdoQuery
Dim del_tmchel As rdoQuery
Dim del_tmchel1 As rdoQuery
Dim del_gtDetail1 As rdoQuery
Dim UPDDINTER_QR As rdoQuery
Dim str As String
Dim REQ1 As String

Private Sub cmd_back_Click()
Me.Hide
DIS_MAIN.show
End Sub

Private Sub cmd_clickit_Click()
'On Error GoTo MSG

    UPDRREQ_QR(0) = REQ1
    UPDRREQ_QR.Execute
    
    del_micro(0) = REQ1
    del_micro.Execute
    
    DEL_COMMENT(0) = REQ1
    DEL_COMMENT.Execute
    
    del_tmchel(0) = REQ1
    del_tmchel.Execute
    
    del_tmchel1(0) = REQ1
    del_tmchel1.Execute
    
    UPDDINTER_QR(0) = REQ1
    UPDDINTER_QR.Execute
    
    
    del_gtDetail1(0) = REQ1
    del_gtDetail1.Execute
    
    MsgBox ("Done......!")
    txt_reqno.Text = ""
    
    
    'GTDETAILS
    Exit Sub
MSG:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Load()

logonfrm.Main
datetime
str = "UPDATE RISLIB.RREQUISITN SET REQNSTAT='E',COMMITFLAG='' WHERE REQNNO=?"
Set UPDRREQ_QR = rdb.CreateQuery("UPDATE111", str)

'str = "DELETE  FROM DISLIB.DMPROCEDUR WHERE REQNO=?"
'Set del_micro = rdb.CreateQuery("DELETE11", str)

str = "DELETE FROM DISLIB.DMBCCOMM WHERE REQNO=?"
Set DEL_COMMENT = rdb.CreateQuery("DEL_COMMENT", str)

str = "DELETE  FROM DISLIB.DIMBIOCHEL WHERE REQNO=?"
Set del_tmchel = rdb.CreateQuery("", str)

str = "DELETE  FROM DISLIB.DIMBIOCHE1 WHERE REQNO=?"
Set del_tmchel1 = rdb.CreateQuery("", str)


str = "UPDATE INTERFACE.DINTERFACE SET IFLAG=2 WHERE REQNO=? AND IFLAG=3"
Set UPDDINTER_QR = rdb.CreateQuery("", str)


str = "DELETE FROM DISLIB.GTDETAILS WHERE REQNO=?"
Set del_gtDetail1 = rdb.CreateQuery("", str)

'TXT_REQNO.SetFocus
txt_reqno.Text = ""

End Sub

Private Sub txt_reqno_LostFocus()
REQ1 = ""
REQ1 = UCase(Trim(txt_reqno.Text))
txt_reqno.Text = REQ1
End Sub
